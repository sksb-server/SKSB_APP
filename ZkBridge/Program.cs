namespace SKSB_App.ZkBridge;

using System.Net.Sockets;
using System.Text;
using System.Text.Json;

public static class Program
{
    const ushort CMD_CONNECT = 1000, CMD_EXIT = 1001, CMD_AUTH = 1102, CMD_ATTLOG_RRQ = 13;
    const ushort CMD_USERTEMP_RRQ = 9, CMD_USER_WRQ = 8, CMD_REFRESH_DATA = 1013;
    const ushort CMD_ACK_OK = 2000, CMD_ACK_UNAUTH = 2005, CMD_PREPARE_DATA = 1500, CMD_DATA = 1501, CMD_FREE_DATA = 1502;
    const int TimeoutMs = 10000;

    public static int Main(string[] args)
    {
        if (args.Length >= 4 && int.TryParse(args[1], out int p2) && int.TryParse(args[2], out int k2))
        {
            try
            {
                if (args[3] == "users" && args.Length == 4)
                {
                    var us = new Client(args[0], p2, k2).FetchUsers()
                        .Select(u => new { u.Uid, u.EnrollNumber, u.Name, u.Role, Card = u.Card });
                    Console.Out.Write(JsonSerializer.Serialize(us));
                    return 0;
                }
                if (args[3] == "remapuser" && (args.Length == 6 || args.Length == 7) && int.TryParse(args[4], out int ruid))
                {
                    var (en, nm2) = new Client(args[0], p2, k2).RemapUser(ruid, args[5], args.Length == 7 ? args[6] : null);
                    Console.Out.Write(JsonSerializer.Serialize(new { Ok = true, Uid = ruid, EnrollNumber = en, Name = nm2 }));
                    return 0;
                }
                if (args[3] == "applyusers" && args.Length == 4)
                {
                    var updates = JsonSerializer.Deserialize<List<UserUpdate>>(Console.In.ReadToEnd(), new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                                  ?? throw new IOException("No updates supplied on stdin.");
                    int count = new Client(args[0], p2, k2).ApplyUsers(updates);
                    Console.Out.Write(JsonSerializer.Serialize(new { Ok = true, Updated = count }));
                    return 0;
                }
                if (args[3] == "setuser" && args.Length == 6)
                {
                    var (nm, created) = new Client(args[0], p2, k2).SetUser(args[4], args[5]);
                    Console.Out.Write(JsonSerializer.Serialize(new { Ok = true, EnrollNumber = args[4], Name = nm, Created = created }));
                    return 0;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 2;
            }
        }
        if (args.Length < 4 || args[3] != "attlog" || args.Length > 5
            || !int.TryParse(args[1], out int port) || !int.TryParse(args[2], out int commKey))
        {
            Console.Error.WriteLine("Usage: ZkBridge.exe <ip> <port> <commKey> attlog [yyyy-MM-dd] | users | setuser <enrollNumber> <name> | remapuser <uid> <newEnrollNumber> [newName] | applyusers (JSON array on stdin)");
            return 1;
        }
        DateOnly? since = null;
        if (args.Length == 5)
        {
            if (!DateOnly.TryParseExact(args[4], "yyyy-MM-dd", out var d))
            {
                Console.Error.WriteLine("Invalid date, expected yyyy-MM-dd.");
                return 1;
            }
            since = d;
        }

        try
        {
            var records = new Client(args[0], port, commKey).FetchAttendance();
            if (since.HasValue) records = records.Where(r => DateOnly.FromDateTime(r.Timestamp) >= since.Value).ToList();
            var output = records.Select(r => new { r.EnrollNumber, Timestamp = r.Timestamp.ToString("yyyy-MM-ddTHH:mm:ss") });
            Console.Out.Write(JsonSerializer.Serialize(output));
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 2;
        }
    }

    sealed record Punch(string EnrollNumber, DateTime Timestamp);
    sealed record UserUpdate(int Uid, string EnrollNumber, string? Name = null);

    sealed class Client
    {
        readonly string _ip; readonly int _port; readonly int _commKey;
        NetworkStream _s = null!;
        ushort _session, _reply;

        public Client(string ip, int port, int commKey) { _ip = ip; _port = port; _commKey = commKey; }

        TcpClient? _tcp;

        void ConnectAndAuth()
        {
            _tcp = new TcpClient { ReceiveTimeout = TimeoutMs, SendTimeout = TimeoutMs };
            if (!_tcp.ConnectAsync(_ip, _port).Wait(TimeoutMs) || !_tcp.Connected)
                throw new IOException($"Timed out connecting to {_ip}:{_port}.");
            _s = _tcp.GetStream();
            _s.ReadTimeout = TimeoutMs; _s.WriteTimeout = TimeoutMs;

            var (cmd, _) = Send(CMD_CONNECT, Array.Empty<byte>());
            if (cmd == CMD_ACK_UNAUTH || (cmd == CMD_ACK_OK && _commKey != 0))
            {
                (cmd, _) = Send(CMD_AUTH, MakeCommKey(_commKey, _session));
                if (cmd != CMD_ACK_OK) throw new IOException("Authentication failed (check comm key).");
            }
            else if (cmd != CMD_ACK_OK) throw new IOException($"Connect rejected (code {cmd}).");
        }

        void Disconnect()
        {
            try { Send(CMD_EXIT, Array.Empty<byte>()); } catch { }
            _tcp?.Dispose();
        }

        public List<Punch> FetchAttendance()
        {
            ConnectAndAuth();
            try { return Decode(ReadDataBlock(CMD_ATTLOG_RRQ, Array.Empty<byte>())); }
            finally { Disconnect(); }
        }

        public List<ZkUser> FetchUsers()
        {
            ConnectAndAuth();
            try { return DecodeUsers(ReadUserBlock(), out _); }
            finally { Disconnect(); }
        }

        byte[] ReadUserBlock() => ReadDataBlock(CMD_USERTEMP_RRQ, new byte[] { 0x05 });

        public (string Name, bool Created) SetUser(string enroll, string name)
        {
            ConnectAndAuth();
            try
            {
                var raw = ReadUserBlock();
                var users = DecodeUsers(raw, out int recSize);
                if (recSize == 0) recSize = 72;
                var existing = users.FirstOrDefault(u => u.EnrollNumber == enroll);
                byte[] rec;
                bool created = existing == null;
                if (existing != null) rec = (byte[])existing.Raw.Clone();
                else
                {
                    rec = new byte[recSize];
                    int maxUid = users.Count == 0 ? 0 : users.Max(u => u.Uid);
                    BitConverter.GetBytes((ushort)(maxUid + 1)).CopyTo(rec, 0);
                    rec[2] = 0;
                    WriteEnroll(rec, recSize, enroll);
                }
                string written = WriteName(rec, recSize, name);

                var (c1, _) = Send(CMD_USER_WRQ, rec);
                if (c1 != CMD_ACK_OK) throw new IOException($"Set user rejected (code {c1}).");
                var (c2, _) = Send(CMD_REFRESH_DATA, Array.Empty<byte>());
                if (c2 != CMD_ACK_OK) throw new IOException($"Refresh rejected (code {c2}).");
                return (written, created);
            }
            finally { Disconnect(); }
        }

        static string WriteName(byte[] rec, int recSize, string name)
        {
            int nameOff = recSize == 72 ? 11 : 8, nameLen = recSize == 72 ? 24 : 8;
            var nb = new byte[nameLen];
            var enc = Encoding.ASCII.GetBytes(new string(name.Select(c => c < 32 || c > 126 ? '?' : c).ToArray()));
            Array.Copy(enc, nb, Math.Min(enc.Length, nameLen - 1));
            Array.Copy(nb, 0, rec, nameOff, nameLen);
            return Encoding.ASCII.GetString(nb).TrimEnd('\0');
        }

        static void WriteEnroll(byte[] rec, int recSize, string enroll)
        {
            if (recSize == 72)
            {
                var eb = Encoding.ASCII.GetBytes(enroll);
                if (eb.Length > 23) throw new IOException("EnrollNumber too long (max 23 characters).");
                Array.Clear(rec, 48, 24);
                eb.CopyTo(rec, 48);
            }
            else
            {
                if (!uint.TryParse(enroll, out uint n)) throw new IOException("EnrollNumber must be numeric for this device.");
                BitConverter.GetBytes(n).CopyTo(rec, 24);
            }
        }

        void WriteUser(byte[] rec)
        {
            var (c, _) = Send(CMD_USER_WRQ, rec);
            if (c != CMD_ACK_OK) throw new IOException($"Set user rejected (code {c}).");
        }

        void Refresh()
        {
            var (c, _) = Send(CMD_REFRESH_DATA, Array.Empty<byte>());
            if (c != CMD_ACK_OK) throw new IOException($"Refresh rejected (code {c}).");
        }

        public (string EnrollNumber, string Name) RemapUser(int uid, string newEnroll, string? newName)
        {
            ConnectAndAuth();
            try
            {
                var users = DecodeUsers(ReadUserBlock(), out int recSize);
                var existing = users.FirstOrDefault(u => u.Uid == uid) ?? throw new IOException($"User with Uid {uid} not found on device.");
                var clash = users.FirstOrDefault(u => u.Uid != uid && u.EnrollNumber == newEnroll);
                if (clash != null) throw new IOException($"EnrollNumber '{newEnroll}' is already used by Uid {clash.Uid}.");
                var rec = (byte[])existing.Raw.Clone();
                WriteEnroll(rec, recSize, newEnroll);
                string name = newName == null ? existing.Name : WriteName(rec, recSize, newName);
                WriteUser(rec);
                Refresh();
                return (newEnroll, name);
            }
            finally { Disconnect(); }
        }

        public int ApplyUsers(List<UserUpdate> updates)
        {
            ConnectAndAuth();
            try
            {
                var users = DecodeUsers(ReadUserBlock(), out int recSize);
                if (updates.Select(u => u.Uid).Distinct().Count() != updates.Count) throw new IOException("Duplicate Uid in updates.");
                var final = users.ToDictionary(u => u.Uid, u => u.EnrollNumber);
                var recs = new List<byte[]>();
                foreach (var up in updates)
                {
                    var existing = users.FirstOrDefault(u => u.Uid == up.Uid) ?? throw new IOException($"User with Uid {up.Uid} not found on device.");
                    if (string.IsNullOrEmpty(up.EnrollNumber)) throw new IOException($"Empty EnrollNumber for Uid {up.Uid}.");
                    var rec = (byte[])existing.Raw.Clone();
                    WriteEnroll(rec, recSize, up.EnrollNumber);
                    if (up.Name != null) WriteName(rec, recSize, up.Name);
                    final[up.Uid] = up.EnrollNumber;
                    recs.Add(rec);
                }
                var dup = final.GroupBy(kv => kv.Value).FirstOrDefault(g => g.Count() > 1);
                if (dup != null) throw new IOException($"Resulting EnrollNumber '{dup.Key}' would be duplicated (Uids {string.Join(", ", dup.Select(k => k.Key))}).");
                foreach (var rec in recs) WriteUser(rec);
                Refresh();
                return recs.Count;
            }
            finally { Disconnect(); }
        }

        byte[] ReadDataBlock(ushort reqCmd, byte[] reqPayload)
        {
            var (cmd, payload) = Send(reqCmd, reqPayload);
            if (cmd == CMD_ACK_OK) return Array.Empty<byte>();
            if (cmd == CMD_DATA) return payload;
            if (cmd != CMD_PREPARE_DATA || payload.Length < 4)
                throw new IOException($"Unexpected response to attendance request (code {cmd}).");

            int total = BitConverter.ToInt32(payload, 0);
            using var ms = new MemoryStream();
            while (ms.Length < total)
            {
                var (c, p) = ReadPacket();
                if (c == CMD_DATA) ms.Write(p);
                else if (c == CMD_ACK_OK) break;
                else throw new IOException($"Unexpected packet while reading data (code {c}).");
            }
            // trailing ACK_OK, if any
            if (ms.Length >= total)
            {
                try { _s.ReadTimeout = 1500; ReadPacket(); } catch { } finally { _s.ReadTimeout = TimeoutMs; }
            }
            try { Send(CMD_FREE_DATA, Array.Empty<byte>()); } catch { }
            return ms.ToArray();
        }

        (ushort cmd, byte[] payload) Send(ushort cmd, byte[] data)
        {
            var pkt = new byte[8 + data.Length];
            BitConverter.GetBytes(cmd).CopyTo(pkt, 0);
            BitConverter.GetBytes(_session).CopyTo(pkt, 4);
            BitConverter.GetBytes(_reply).CopyTo(pkt, 6);
            data.CopyTo(pkt, 8);
            BitConverter.GetBytes(Checksum(pkt)).CopyTo(pkt, 2);

            var frame = new byte[8 + pkt.Length];
            frame[0] = 0x50; frame[1] = 0x50; frame[2] = 0x82; frame[3] = 0x7D;
            BitConverter.GetBytes(pkt.Length).CopyTo(frame, 4);
            pkt.CopyTo(frame, 8);
            _s.Write(frame, 0, frame.Length);
            _reply = (ushort)(_reply + 1);

            var r = ReadPacket();
            if (cmd == CMD_CONNECT) _session = _lastSession;
            return r;
        }

        ushort _lastSession;

        (ushort cmd, byte[] payload) ReadPacket()
        {
            var hdr = ReadExact(8);
            if (hdr[0] != 0x50 || hdr[1] != 0x50 || hdr[2] != 0x82 || hdr[3] != 0x7D)
                throw new IOException("Invalid ZKTeco frame header.");
            int len = BitConverter.ToInt32(hdr, 4);
            if (len < 8 || len > 64 * 1024 * 1024) throw new IOException("Invalid ZKTeco frame length.");
            var body = ReadExact(len);
            _lastSession = BitConverter.ToUInt16(body, 4);
            return (BitConverter.ToUInt16(body, 0), body[8..]);
        }

        byte[] ReadExact(int n)
        {
            var buf = new byte[n]; int off = 0;
            while (off < n)
            {
                int r = _s.Read(buf, off, n - off);
                if (r <= 0) throw new IOException("Connection closed by device.");
                off += r;
            }
            return buf;
        }
    }

    sealed record ZkUser(int Uid, string EnrollNumber, string Name, int Role, uint Card, byte[] Raw);

    static string CStr(byte[] b, int off, int len) =>
        Encoding.UTF8.GetString(b, off, len).Split('\0')[0].Trim();

    static List<ZkUser> DecodeUsers(byte[] data, out int recSize)
    {
        var list = new List<ZkUser>();
        recSize = 0;
        int off = data.Length >= 4 ? 4 : 0;
        int len = data.Length - off;
        if (len <= 0) return list;
        recSize = len % 72 == 0 ? 72 : len % 28 == 0 ? 28 : 0;
        if (recSize == 0) throw new IOException($"Unrecognized user record layout ({len} bytes).");
        for (; off + recSize <= data.Length; off += recSize)
        {
            var raw = data[off..(off + recSize)];
            int uid = BitConverter.ToUInt16(raw, 0), role = raw[2];
            if (recSize == 72)
                list.Add(new ZkUser(uid, CStr(raw, 48, 24), CStr(raw, 11, 24), role, BitConverter.ToUInt32(raw, 35), raw));
            else
                list.Add(new ZkUser(uid, BitConverter.ToUInt32(raw, 24).ToString(), CStr(raw, 8, 8), role, BitConverter.ToUInt32(raw, 16), raw));
        }
        return list;
    }

    static ushort Checksum(byte[] p)
    {
        // checksum field (bytes 2-3) assumed zero
        uint sum = 0;
        for (int i = 0; i + 1 < p.Length; i += 2) sum += BitConverter.ToUInt16(p, i);
        if (p.Length % 2 == 1) sum += p[^1];
        while (sum >> 16 != 0) sum = (sum & 0xFFFF) + (sum >> 16);
        return (ushort)~sum;
    }

    static byte[] MakeCommKey(int key, ushort session)
    {
        uint k = 0;
        for (int i = 0; i < 32; i++)
            k = (key & (1 << i)) != 0 ? (uint)((k << 1) | 1) : k << 1;
        k += session;
        var b = BitConverter.GetBytes(k);
        b[0] ^= (byte)'Z'; b[1] ^= (byte)'K'; b[2] ^= (byte)'S'; b[3] ^= (byte)'O';
        ushort a = BitConverter.ToUInt16(b, 0), c = BitConverter.ToUInt16(b, 2);
        var o = new byte[4];
        BitConverter.GetBytes(c).CopyTo(o, 0);
        BitConverter.GetBytes(a).CopyTo(o, 2);
        byte B = 0xFF; // ticks
        o[0] ^= B; o[1] ^= B; o[2] = B; o[3] ^= B;
        return o;
    }

    static DateTime DecodeTime(uint t)
    {
        int sec = (int)(t % 60), min = (int)(t / 60 % 60), hour = (int)(t / 3600 % 24);
        int day = (int)(t / 86400 % 31) + 1, month = (int)(t / 2678400 % 12) + 1, year = (int)(t / 32140800) + 2000;
        return new DateTime(year, month, Math.Min(day, DateTime.DaysInMonth(year, month)), hour, min, sec);
    }

    static List<Punch> Decode(byte[] buf)
    {
        var list = new List<Punch>();
        int off = 0;
        if (buf.Length >= 4)
        {
            int size = BitConverter.ToInt32(buf, 0);
            if (size == buf.Length - 4) off = 4;
        }
        int len = buf.Length - off;
        if (len <= 0) return list;
        int rec = len % 40 == 0 ? 40 : len % 16 == 0 ? 16 : len % 8 == 0 ? 8 : 40;

        for (; off + rec <= buf.Length; off += rec)
        {
            try
            {
                string id; uint t;
                if (rec == 40)
                {
                    id = Encoding.ASCII.GetString(buf, off + 2, 24).Split('\0')[0].Trim();
                    t = BitConverter.ToUInt32(buf, off + 27);
                }
                else if (rec == 16)
                {
                    id = BitConverter.ToUInt32(buf, off).ToString();
                    t = BitConverter.ToUInt32(buf, off + 4);
                }
                else
                {
                    id = BitConverter.ToUInt16(buf, off).ToString();
                    t = BitConverter.ToUInt32(buf, off + 4);
                }
                if (id.Length == 0) continue;
                list.Add(new Punch(id, DecodeTime(t)));
            }
            catch (ArgumentOutOfRangeException) { }
        }
        return list;
    }
}
