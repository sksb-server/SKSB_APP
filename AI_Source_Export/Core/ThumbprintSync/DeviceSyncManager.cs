using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace SKSB_App.ThumbprintSync
{
    public class TerminalUser
    {
        public string EnrollNumber { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public int Privilege { get; set; }
        public string CardNumber { get; set; } = string.Empty;
    }

    [ComVisible(true)]
    public class DeviceSyncController : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabAllocation' label='Work Allocation'>
                    <group id='grpDevice' label='Terminal Direct Sync'>
                      <button id='btnImportTerminal' 
                              label='Pull Users From Reader' 
                              size='large' 
                              imageMso='ContactAddNew' 
                              onAction='OnPullUsersClicked' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnPullUsersClicked(IRibbonControl control)
        {
            Task.Run(async () =>
            {
                const string ip = "192.168.0.110";
                const int port = 4370;

                try
                {
                    List<TerminalUser> users = await FetchEnrolledUsersAsync(ip, port);

                    if (users == null || users.Count == 0)
                    {
                        MessageBox.Show($"Connected to reader {ip}, but 0 enrolled user records were parsed from memory.",
                                        "Reader Query Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Dispatch to Excel main thread
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            PopulateEmployeesTable(users);
                            MessageBox.Show($"Successfully imported {users.Count} employees directly from reader ({ip}).",
                                            "Sync Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception renderEx)
                        {
                            MessageBox.Show($"Error writing to sheet: {renderEx.Message}\n\n{renderEx.StackTrace}",
                                            "Excel Render Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Connection failed:\n\n{ex.Message}\n\n({ex.GetType().Name})",
                                    "Terminal Communication Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });
        }

        private async Task<List<TerminalUser>> private async Task<List<TerminalUser>> FetchEnrolledUsersAsync(string ip, int port, int commKey = 0)
        {
            return await Task.Run(() =>
            {
                // Locate ZkBridge.exe relative to the running XLL add-in
                string addinDirectory = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
                string bridgeExePath = Path.Combine(addinDirectory, "ZkBridge.exe");

                if (!File.Exists(bridgeExePath))
                {
                    // Fallback check current directory
                    bridgeExePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ZkBridge.exe");
                    if (!File.Exists(bridgeExePath))
                    {
                        throw new FileNotFoundException($"Cannot locate ZkBridge.exe at:\n{bridgeExePath}");
                    }
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = bridgeExePath,
                    Arguments = $"{ip} {port} {commKey}",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit(15000);

                    if (process.ExitCode != 0 || !string.IsNullOrWhiteSpace(error))
                    {
                        throw new Exception($"Bridge Error (Code {process.ExitCode}): {error}");
                    }

                    if (string.IsNullOrWhiteSpace(output))
                    {
                        return new List<TerminalUser>();
                    }

                    var serializer = new JavaScriptSerializer();
                    return serializer.Deserialize<List<TerminalUser>>(output);
                }
            });
        }

        private byte[] CreateZkPacket(ushort command, ushort sessionId, ushort replyNumber, byte[] data)
        {
            byte[] packet = new byte[8 + data.Length];
            BitConverter.GetBytes(command).CopyTo(packet, 0);
            BitConverter.GetBytes((ushort)0).CopyTo(packet, 2);
            BitConverter.GetBytes(sessionId).CopyTo(packet, 4);
            BitConverter.GetBytes(replyNumber).CopyTo(packet, 6);
            if (data.Length > 0) data.CopyTo(packet, 8);

            int chk = 0;
            for (int i = 0; i < packet.Length; i += 2)
            {
                if (i == 2) continue;
                ushort val = (i + 1 < packet.Length) ? BitConverter.ToUInt16(packet, i) : packet[i];
                chk += val;
                if (chk > 0xFFFF) chk = (chk & 0xFFFF) + 1;
            }
            ushort checksum = (ushort)(~chk & 0xFFFF);
            BitConverter.GetBytes(checksum).CopyTo(packet, 2);

            return packet;
        }

        private byte[] MakeAuthPayload(int commKey, ushort sessionId)
        {
            int k = 0;
            for (int i = 0; i < 32; i++)
            {
                if ((commKey & (1 << i)) > 0)
                    k = (k << 1 | 1);
                else
                    k = (k << 1);
            }
            k += sessionId;

            byte[] keyBytes = BitConverter.GetBytes(k);
            byte[] encrypted = new byte[4];
            encrypted[0] = (byte)(keyBytes[0] ^ 0x50);
            encrypted[1] = (byte)(keyBytes[1] ^ 0x4F);
            encrypted[2] = (byte)(keyBytes[2] ^ 0x53);
            encrypted[3] = (byte)(keyBytes[3] ^ 0x4F);

            return encrypted;
        }

        private void PopulateEmployeesTable(List<TerminalUser> users)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Workbook wb = app.ActiveWorkbook ?? app.Workbooks.Add();

            Excel.Worksheet ws;
            try
            {
                ws = (Excel.Worksheet)wb.Sheets["Employees"];
            }
            catch
            {
                ws = (Excel.Worksheet)wb.Sheets.Add();
                ws.Name = "Employees";
            }

            ws.Cells.Clear();

            string[] headers = { "EnrollNumber", "Employee Name", "Privilege", "Card Number", "Source of Truth" };
            for (int c = 0; c < headers.Length; c++)
            {
                ws.Cells[1, c + 1] = headers[c];
            }

            object[,] data = new object[users.Count, headers.Length];
            for (int i = 0; i < users.Count; i++)
            {
                data[i, 0] = users[i].EnrollNumber;
                data[i, 1] = users[i].Name;
                data[i, 2] = users[i].Privilege == 14 ? "Admin" : "User";
                data[i, 3] = users[i].CardNumber;
                data[i, 4] = "Terminal (192.168.0.110)";
            }

            Excel.Range dataRange = ws.Range[ws.Cells[2, 1], ws.Cells[users.Count + 1, headers.Length]];
            dataRange.Value2 = data;

            Excel.Range fullTableRange = ws.Range[ws.Cells[1, 1], ws.Cells[users.Count + 1, headers.Length]];
            try
            {
                Excel.ListObject table = ws.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    fullTableRange,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes
                );
                table.Name = "Table_Employees";
                table.TableStyle = "TableStyleMedium9";
            }
            catch
            {
                // Table already formatted
            }

            ws.Columns.AutoFit();
            ((Excel.Range)ws.Columns[1]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)ws.Columns[3]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }
    }
}