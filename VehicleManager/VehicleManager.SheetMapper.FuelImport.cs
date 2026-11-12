namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;

internal static partial class SheetMapper
{
    // =========================================================================
    // FUEL STATEMENT IMPORTER (PETRONAS SMARTPAY & SHELL FLEET HUB)
    // =========================================================================
    /// <summary>
    /// Duplicate-detection key. Deliberately excludes CurrentOdometer and any plate reassigned by odometer matching,
    /// so re-importing a statement after "Sort &amp; Match Odometers" is still recognised as a duplicate.
    /// </summary>
    private static string BuildFuelSignature(OdometerLog log)
    {
        string card = (log.FuelCardNumber ?? string.Empty).Trim();
        string id = string.IsNullOrWhiteSpace(card) ? VehiclePlateHelper.NormalizePlate(log.VehiclePlate) : card;
        return $"{log.Date:yyyyMMdd}_{(log.Time ?? string.Empty).Trim()}_{(log.ReceiptNumber ?? string.Empty).Trim()}_{log.Amount:F2}_{log.FuelAmount:F2}_{log.PriceAtFueling:F2}_{id}";
    }

    internal static (int ImportedCount, int SkippedCount, decimal TotalAmount, string ProviderName) ImportFuelStatement(string filePath)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic wb = xlApp.ActiveWorkbook;
        if (wb == null) throw new InvalidOperationException("No active workbook found.");

        dynamic? wsOdo = SheetMapperHelper.ResolveSheetSafely((object)wb, OdometerAliases, "Odometers");
        if (wsOdo == null) throw new InvalidOperationException("Worksheet 'Odometers' was not found in active workbook.");

        List<OdometerLog> existingLogs = TableMapper.ExtractFromTable<OdometerLog>((object)wsOdo, "OdometersTable");
        List<FuelCardRecord> fuelCards = TableMapper.ExtractFromTable<FuelCardRecord>((object)wsOdo, "FuelCardsTable");

        dynamic? wsMaster = SheetMapperHelper.ResolveSheetSafely((object)wb, AllVehicleAliases, MasterSheetName);
        List<VehicleRecord> masterVehicles = wsMaster != null
            ? TableMapper.ExtractFromTable<VehicleRecord>((object)wsMaster, MasterTableName)
            : new List<VehicleRecord>();

        // 1. Load Staff Directory first so the PDF parser can use known staff names to split Station vs Driver
        var staffList = OfflineSnapshot.Load<List<StaffRecord>>() ?? new List<StaffRecord>();
        var staffExactMap = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        var staffShortNameMap = new Dictionary<long, string>();
        var knownDriverNames = new List<string>();

        foreach (var s in staffList)
        {
            if (s.StaffID <= 0) continue;
            string fn = NormalizeName(s.StaffDetails?.StaffName);
            string sn = NormalizeName(s.StaffDetails?.StaffShortName);
            string resolvedShort = !string.IsNullOrEmpty(sn) ? sn : fn;

            if (!string.IsNullOrEmpty(sn))
            {
                staffExactMap[sn] = s.StaffID;
                knownDriverNames.Add(sn);
            }
            if (!string.IsNullOrEmpty(fn))
            {
                if (!staffExactMap.ContainsKey(fn)) staffExactMap[fn] = s.StaffID;
                knownDriverNames.Add(fn);
            }

            staffShortNameMap[s.StaffID] = resolvedShort;
        }

        // 2. Build Fuel Card Registry (Normalized digits -> FuelCardRecord)
        var cardMap = new Dictionary<string, FuelCardRecord>(StringComparer.OrdinalIgnoreCase);
        foreach (var c in fuelCards)
        {
            string normCard = SanitizeCardToken(c.FuelCardNumber);
            if (!string.IsNullOrWhiteSpace(normCard))
            {
                cardMap[normCard] = c;
            }
        }

        // 3. Parse Statement (passing knownDriverNames for accurate Petronas STM Station/Driver splitting)
        var parsedItems = FuelStatementParser.Parse(filePath, knownDriverNames, out string providerName);

        if (parsedItems.Count == 0)
        {
            AppLogger.Warn($"[FuelImport] Import yielded 0 records for {Path.GetFileName(filePath)}.");
            return (0, 0, 0m, providerName);
        }

        var vehicleTypeMap = masterVehicles
            .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
            .GroupBy(v => VehiclePlateHelper.NormalizePlate(v.VehiclePlate), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().VehicleDetail.Type.ToString(), StringComparer.OrdinalIgnoreCase);

        long maxTxId = 0;
        var existingSignatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var ex in existingLogs)
        {
            long id = GetProp<long>(ex, "FuelTransactionID", "TransactionID", "OdometerID");
            if (id > maxTxId) maxTxId = id;
            if (ex.Date != default && ex.Amount > 0)
                existingSignatures.Add(BuildFuelSignature(ex));
        }

        // 1. Build candidate log objects from statement
        var candidateLogs = new List<OdometerLog>();
        foreach (var raw in parsedItems)
        {
            string resolvedCard = ResolveFullCardNumber(raw.CardNumber, cardMap, out FuelCardRecord? matchedCard);
            string plate = VehiclePlateHelper.FormatPlate(raw.VehiclePlate);

            if ((string.IsNullOrWhiteSpace(plate) || plate.Equals("SPARE", StringComparison.OrdinalIgnoreCase)) &&
                matchedCard != null && !string.IsNullOrWhiteSpace(matchedCard.CardOwnerID) &&
                !long.TryParse(matchedCard.CardOwnerID.Trim(), out _))
            {
                plate = VehiclePlateHelper.FormatPlate(matchedCard.CardOwnerID);
            }

            string vType = raw.VehicleType;
            if (string.IsNullOrWhiteSpace(vType) && vehicleTypeMap.TryGetValue(VehiclePlateHelper.NormalizePlate(plate), out var vt))
            {
                vType = vt;
            }

            long staffId = raw.StaffID;
            string staffName = raw.StaffName;
            if (staffId <= 0 && !string.IsNullOrWhiteSpace(staffName))
                staffId = ResolveStaffId(staffName, staffExactMap, staffList);

            if (staffId <= 0 && matchedCard != null)
            {
                if (!string.IsNullOrWhiteSpace(matchedCard.CardOwnerID) &&
                    long.TryParse(matchedCard.CardOwnerID.Trim(), out long cardOwnerStaffId) &&
                    cardOwnerStaffId > 0)
                {
                    staffId = cardOwnerStaffId;
                }
                else if (!string.IsNullOrWhiteSpace(matchedCard.CardOwnerName))
                {
                    staffId = ResolveStaffId(matchedCard.CardOwnerName, staffExactMap, staffList);
                    if (string.IsNullOrWhiteSpace(staffName)) staffName = matchedCard.CardOwnerName.Trim();
                }
            }

            if (staffId > 0 && staffShortNameMap.TryGetValue(staffId, out var canonicalShortName))
                staffName = canonicalShortName;

            var log = new OdometerLog
            {
                Date = raw.Date,
                Time = raw.Time?.Trim() ?? string.Empty,
                VehiclePlate = plate,
                CurrentOdometer = raw.Odometer > 0 ? (decimal)raw.Odometer : 0m,
                FuelCardNumber = resolvedCard,
                StationName = raw.StationName ?? string.Empty,
                FuelType = raw.FuelType ?? string.Empty,
                PriceAtFueling = raw.PriceAtFueling,
                FuelAmount = raw.FuelAmount,
                Amount = raw.Amount,
                ReceiptNumber = raw.ReceiptNumber ?? string.Empty
            };
            if (!string.IsNullOrWhiteSpace(vType)) SetProp(log, "VehicleType", vType);
            if (staffId > 0) log.StaffID = staffId;
            if (!string.IsNullOrWhiteSpace(staffName)) log.StaffName = staffName;

            candidateLogs.Add(log);
        }

        // 2. Filter out duplicates while retaining clean non-duplicate records
        var cleanTransactions = TableRowManager.FilterDuplicates(
            incomingRecords: candidateLogs,
            keySelector: BuildFuelSignature,
            displaySelector: log =>
            {
                DateOnly d = GetProp<DateOnly>(log, "Date");
                string pl = GetProp<string>(log, "VehiclePlate") ?? string.Empty;
                decimal amt = GetProp<decimal>(log, "Amount", "TransactionAmount");
                return $"{d:dd/MM/yyyy} - {pl} (RM {amt:N2})";
            },
            entityLabel: "Fuel Transaction",
            duplicateWarnings: out var duplicateWarnings,
            existingRecords: existingLogs
        );

        // Cancelled by user in dialog
        if (cleanTransactions == null)
        {
            SheetOperations.ShowStatusConfirmation("Fuel import cancelled by user.");
            return (0, 0, 0m, providerName);
        }

        int skippedCount = candidateLogs.Count - cleanTransactions.Count;
        decimal totalAmount = 0m;

        // 3. Assign IDs to clean records and append to OdometersTable
        foreach (var log in cleanTransactions)
        {
            maxTxId++;
            log.OdometerID = maxTxId;
            totalAmount += GetProp<decimal>(log, "Amount", "TransactionAmount");
        }

        if (cleanTransactions.Count > 0)
        {
            var combined = new List<OdometerLog>(existingLogs.Count + cleanTransactions.Count);
            combined.AddRange(existingLogs);
            combined.AddRange(cleanTransactions);

            // Normalise hour-meters, resolve wrong/unrecognised plates, zero dummy readings, sort and chain OdometerBefore.
            var reconcile = ReconcileAndSortOdometerLogs(combined, masterVehicles, cardMap);
            AppLogger.Info($"[FuelImport] Reconcile: matched={reconcile.VehiclesMatched}, hours={reconcile.HoursNormalized}, invalid={reconcile.InvalidZeroed}, flagged={reconcile.FlaggedPlates}");

            WriteBackOdometers(xlApp, wsOdo, wsMaster, combined, masterVehicles);
        }

        return (cleanTransactions.Count, skippedCount, totalAmount, providerName);
    }

    // =========================================================================
    // ODOMETER RECONCILIATION ENGINE (shared by Import Fuel Statement & Sort & Match Odometers)
    // =========================================================================
    internal sealed record OdometerReconcileResult(
        int RowsSorted, int VehiclesMatched, int HoursNormalized, int InvalidZeroed, int FlaggedPlates);

    private sealed class OdoWork
    {
        public OdometerLog Log = null!;
        public decimal Raw;                 // working reading as typed (or recovered from an [Invalid Odometer] tag)
        public decimal Norm;                // reading after hour-meter normalisation
        public decimal JunkRaw;             // value used for junk tests (tag raw for already-scaled rows)
        public string OrigPlate = string.Empty;
        public string Key = string.Empty;   // group key (normalised preliminary plate)
        public bool AlreadyScaled;          // hours already divided by 10 on a previous run (or typed with a decimal)
        public bool PlateLocked;            // plate was matched on a previous run or hand-edited: never reassign
        public bool Invalid;
        public bool Divided;
        public bool NeedsMatch;
        public bool CanReject;              // no match found => zero the reading
        public bool Matched;
        public bool Ambiguous;
    }

    private sealed class OdoCandidate
    {
        public string Plate = string.Empty;
        public string Key = string.Empty;
        public VehicleType Type = VehicleType.None;
        public bool Hour;
        public decimal MasterOdo;
        public List<(DateTime Stamp, decimal Odo)> Anchors = new();
    }

    // System-generated remark tags (Step 0: state recovery)
    private static readonly Regex ScaledHoursRx = new(@"\[(?<hrs>[\d\.]+)\s*hrs\s*\(raw\s*(?<raw>\d+)\)\]", RegexOptions.Compiled);
    private static readonly Regex MatchedPlateRx = new(@"\[Card\s+(?<card>.+?)\s*->\s*Matched\s+(?<plate>.+?)\s+by Odometer\]", RegexOptions.Compiled);
    private static readonly Regex InvalidOdoRx = new(@"\[Invalid Odometer:\s*(?:repeated\s*)?raw\s*(?<raw>[\d\.]+)[^\]]*\]", RegexOptions.Compiled);
    private static readonly Regex AmbiguousOdoRx = new(@"\[Ambiguous Odometer\s+(?<raw>[\d\.]+):[^\]]+\]", RegexOptions.Compiled);
    private static readonly Regex UnrecognizedPlateRx = new(@"\[Unrecognized Plate:\s*(?<plate>[^\s\]]+)[^\]]*\]", RegexOptions.Compiled);

    private static readonly HashSet<string> PlaceholderPlates =
        new(StringComparer.OrdinalIgnoreCase) { "", "SPARE", "CHK", "UNKNOWN", "NA", "N/A", "NIL" };

    /// <summary>Step 0: reads previous-run tags off Remarks, strips stale error tags, and recovers zeroed readings.</summary>
    private static OdoWork ParseOdometerState(OdometerLog l)
    {
        var w = new OdoWork { Log = l, OrigPlate = (l.VehiclePlate ?? string.Empty).Trim() };
        string rem = l.Remarks ?? string.Empty;
        decimal current = l.CurrentOdometer;
        decimal candidate = 0m;

        var inv = InvalidOdoRx.Match(rem);
        if (inv.Success && decimal.TryParse(inv.Groups["raw"].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var cr))
            candidate = cr;
        rem = InvalidOdoRx.Replace(rem, string.Empty);
        rem = AmbiguousOdoRx.Replace(rem, string.Empty);
        rem = UnrecognizedPlateRx.Replace(rem, string.Empty);

        var sm = ScaledHoursRx.Match(rem);
        if (sm.Success)
        {
            if (current > 0m
                && decimal.TryParse(sm.Groups["hrs"].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var tagHrs)
                && Math.Abs(tagHrs - current) <= 0.05m)
            {
                w.AlreadyScaled = true;
                if (decimal.TryParse(sm.Groups["raw"].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var tagRaw))
                    w.JunkRaw = tagRaw;
            }
            else
            {
                rem = ScaledHoursRx.Replace(rem, string.Empty);   // reading was edited after scaling: tag is stale
            }
        }

        var mm = MatchedPlateRx.Match(rem);
        if (mm.Success)
        {
            w.PlateLocked = true;
            if (VehiclePlateHelper.NormalizePlate(l.VehiclePlate) != VehiclePlateHelper.NormalizePlate(mm.Groups["plate"].Value))
                rem = MatchedPlateRx.Replace(rem, string.Empty);   // user chose a different plate: respect it
        }

        if (!w.AlreadyScaled && current % 1m != 0m) w.AlreadyScaled = true;   // hand-typed decimal hours

        l.Remarks = Regex.Replace(rem, @"\s{2,}", " ").Trim();
        w.Raw = current > 0m ? current : candidate;
        w.Norm = w.Raw;
        if (w.JunkRaw == 0m) w.JunkRaw = w.Raw;
        return w;
    }

    private static void AppendRemarkTag(OdometerLog log, string tag)
    {
        string r = (log.Remarks ?? string.Empty).Trim();
        if (r.Contains(tag, StringComparison.Ordinal)) return;
        log.Remarks = r.Length == 0 ? tag : r + " " + tag;
    }

    private static DateTime OdoStamp(OdometerLog l)
    {
        var t = TimeOnly.MinValue;
        if (!string.IsNullOrWhiteSpace(l.Time))
            TimeOnly.TryParse(l.Time.Trim(), CultureInfo.InvariantCulture, DateTimeStyles.None, out t);
        return l.Date.ToDateTime(t);
    }

    private static decimal MedianOf(List<decimal> values)
    {
        if (values.Count == 0) return 0m;
        var s = values.OrderBy(x => x).ToList();
        int m = s.Count / 2;
        return s.Count % 2 == 1 ? s[m] : (s[m - 1] + s[m]) / 2m;
    }

    private static string F1(decimal v) => v.ToString("F1", CultureInfo.InvariantCulture);
    private static string F0(decimal v) => v.ToString("0", CultureInfo.InvariantCulture);
    private static string RawText(decimal v) => v % 1m == 0m ? F0(v) : F1(v);
    private static string PlateTag(string? p) => string.IsNullOrWhiteSpace(p) ? "BLANK" : VehiclePlateHelper.NormalizePlate(p);

    /// <summary>Step A: dummy / mashed / impossible readings.</summary>
    private static bool IsJunkOdometer(decimal raw)
    {
        if (raw < 10m || raw > 1_500_000m) return true;
        if (raw % 1m != 0m) return false;
        string d = ((long)raw).ToString(CultureInfo.InvariantCulture);
        if (d.Length >= 4 && d.Distinct().Count() == 1) return true;   // 111111, 9999
        if (d.Length >= 5)
        {
            bool asc = true, desc = true;
            for (int i = 1; i < d.Length; i++)
            {
                if (d[i] != d[i - 1] + 1) asc = false;
                if (d[i] != d[i - 1] - 1) desc = false;
            }
            if (asc || desc) return true;                               // 12345, 987654
        }
        return false;
    }

    private static bool IsDieselOnlyRoadType(VehicleType t) =>
        t is VehicleType.Lorry or VehicleType.Tipper_Lorry or VehicleType.Roro_Lorry
          or VehicleType.Crane_Lorry or VehicleType.Skylift;

    private static bool FuelCompatible(OdometerLog l, OdoCandidate c)
    {
        // Subsidised diesel is only sold against road diesel vehicles - never petrol cars or unsubsidised machinery.
        if (l.PriceAtFueling > 0m && l.PriceAtFueling <= 2.50m)
            return !c.Hour && c.Type is not (VehicleType.Car or VehicleType.Motorcycle);

        string ft = l.FuelType ?? string.Empty;
        bool petrol = ft.Contains("RON", StringComparison.OrdinalIgnoreCase)
                   || ft.Contains("PRIMAX", StringComparison.OrdinalIgnoreCase)
                   || ft.Contains("V-Power", StringComparison.OrdinalIgnoreCase);
        return !(petrol && (c.Hour || IsDieselOnlyRoadType(c.Type)));
    }

    /// <summary>
    /// Remark-aware reconciliation: sorts, normalises, matches and chains odometer logs in place.
    /// Idempotent - previous-run tags are parsed first, so no edit is ever applied twice.
    /// </summary>
    internal static OdometerReconcileResult ReconcileAndSortOdometerLogs(
        List<OdometerLog> logs,
        IList<VehicleRecord> masters,
        IDictionary<string, FuelCardRecord>? cardMap = null)
    {
        int matchedCount = 0, hoursNormalized = 0, invalidZeroed = 0;
        if (logs.Count == 0) return new OdometerReconcileResult(0, 0, 0, 0, 0);

        // ---- Known fleet -----------------------------------------------------------------
        var known = new Dictionary<string, VehicleRecord>(StringComparer.OrdinalIgnoreCase);
        foreach (var v in masters)
        {
            if (string.IsNullOrWhiteSpace(v.VehiclePlate)) continue;
            known.TryAdd(VehiclePlateHelper.NormalizePlate(v.VehiclePlate), v);
        }
        bool haveMaster = known.Count > 0;

        bool IsKnownPlate(string? plate)
        {
            string n = VehiclePlateHelper.NormalizePlate(plate);
            if (PlaceholderPlates.Contains(n)) return false;
            return !haveMaster || known.ContainsKey(n);
        }

        // ---- Step 0: parse previous-run state --------------------------------------------
        var works = logs.Select(ParseOdometerState).ToList();

        // Placeholder plates (SPARE / blank): fall back to the vehicle the fuel card is registered to.
        if (cardMap != null && haveMaster)
        {
            foreach (var w in works)
            {
                if (w.PlateLocked || IsKnownPlate(w.Log.VehiclePlate)) continue;
                if (!PlaceholderPlates.Contains(VehiclePlateHelper.NormalizePlate(w.Log.VehiclePlate))) continue;
                if (!cardMap.TryGetValue(SanitizeCardToken(w.Log.FuelCardNumber), out var card)) continue;
                string owner = (card.CardOwnerID ?? string.Empty).Trim();
                if (owner.Length == 0 || long.TryParse(owner, out _)) continue;
                if (known.TryGetValue(VehiclePlateHelper.NormalizePlate(owner), out var ownerVehicle))
                    w.Log.VehiclePlate = ownerVehicle.VehiclePlate;
            }
        }

        foreach (var w in works) w.Key = VehiclePlateHelper.NormalizePlate(w.Log.VehiclePlate);

        void Invalidate(OdoWork w, string why)
        {
            if (w.Invalid) return;
            w.Invalid = true;
            w.NeedsMatch = false;
            w.Log.CurrentOdometer = 0m;
            AppendRemarkTag(w.Log, $"[Invalid Odometer: {why}]");
            invalidZeroed++;
        }

        // ---- Per-vehicle analysis --------------------------------------------------------
        var candidates = new Dictionary<string, OdoCandidate>(StringComparer.OrdinalIgnoreCase);

        foreach (var g in works.GroupBy(w => w.Key))
        {
            string key = g.Key;
            var rows = g.OrderBy(w => OdoStamp(w.Log)).ThenBy(w => w.Log.OdometerID).ToList();
            bool recognized = IsKnownPlate(key);

            // A) Junk / dummy readings, then lazily repeated values.
            var live = new List<OdoWork>();
            foreach (var w in rows)
            {
                if (IsJunkOdometer(w.JunkRaw)) Invalidate(w, $"raw {RawText(w.JunkRaw)}");
                else live.Add(w);
            }
            for (int i = 0; i < live.Count;)
            {
                int j = i;
                while (j + 1 < live.Count && live[j + 1].JunkRaw == live[i].JunkRaw) j++;
                if (j - i + 1 >= 3 &&
                    live.Skip(i).Take(j - i + 1).Select(w => OdoStamp(w.Log)).Distinct().Count() >= 3)
                {
                    for (int k = i; k <= j; k++) Invalidate(live[k], $"repeated raw {RawText(live[k].JunkRaw)}");
                }
                i = j + 1;
            }
            live = live.Where(w => !w.Invalid).ToList();

            // Unrecognised plate: nothing reliable to compare against - every reading is matched individually.
            if (!recognized)
            {
                foreach (var w in live) if (!w.PlateLocked) w.NeedsMatch = true;
                continue;
            }

            known.TryGetValue(key, out var kv);
            VehicleType vType = kv?.VehicleDetail.Type ?? VehicleType.None;
            if (vType == VehicleType.None)
                vType = rows.Select(r => r.Log.VehicleType).FirstOrDefault(t => t != VehicleType.None);
            bool typeDecided = vType != VehicleType.None && (kv != null || !haveMaster);
            decimal masterOdo = kv?.VehicleDetail.CurrentOdometer ?? 0m;

            // B) Hour-meter detection (unscaled readings only).
            var unscaled = live.Where(w => !w.AlreadyScaled).ToList();
            var ratios = new List<decimal>();
            var deltas = new List<decimal>();
            var rates = new List<decimal>();
            for (int i = 1; i < unscaled.Count; i++)
            {
                decimal d = unscaled[i].Raw - unscaled[i - 1].Raw;
                if (d <= 0m) continue;
                deltas.Add(d);
                if (unscaled[i].Log.FuelAmount > 0m) ratios.Add(d / unscaled[i].Log.FuelAmount);
                double days = (OdoStamp(unscaled[i].Log) - OdoStamp(unscaled[i - 1].Log)).TotalDays;
                if (days >= 1.0) rates.Add(d / (decimal)days);
            }
            bool tenthsConsistent = ratios.Count >= 2 && MedianOf(ratios) is >= 0.8m and <= 4.0m;
            // Implied work rate: > 20 hrs/day unscaled but 0.5 - 18 hrs/day once divided by 10.
            decimal medRate = MedianOf(rates);
            bool rateImpliesTenths = rates.Count >= 2 && medRate > 20m && medRate / 10m is >= 0.5m and <= 18m;

            bool hourVehicle;
            if (typeDecided)
            {
                hourVehicle = VehicleRecord.IsHourMeterType(vType);
            }
            else
            {
                bool unsubsidised = live.Count >= 2 && live.Count(w => w.Log.PriceAtFueling > 3.00m) * 2 >= live.Count;
                bool smallSteps = deltas.Count >= 2 && MedianOf(deltas) is >= 1m and <= 40m
                                  && unscaled.All(w => w.Raw < 25000m);
                hourVehicle = unsubsidised && (tenthsConsistent || smallSteps);
            }

            // C) Normalise readings.
            bool scaleAll = hourVehicle && (tenthsConsistent || rateImpliesTenths);
            foreach (var w in live)
            {
                w.Norm = w.AlreadyScaled ? w.Log.CurrentOdometer : w.Raw;
                if (!hourVehicle || w.AlreadyScaled) continue;

                if (scaleAll)
                {
                    w.Norm = Math.Round(w.Raw / 10m, 1);
                    w.Divided = true;
                    continue;
                }

                var others = live.Where(o => o != w).Select(o => o.AlreadyScaled ? o.Log.CurrentOdometer : o.Raw).ToList();
                decimal baseline = others.Count > 0 ? MedianOf(others) : masterOdo;
                decimal tenth = Math.Round(w.Raw / 10m, 1);
                bool scale = baseline > 0m
                    ? Math.Abs(tenth - baseline) <= 500m && Math.Abs(w.Raw - baseline) > 500m   // 14192 vs 1421..1429
                    : w.Raw > 25000m;
                if (scale) { w.Norm = tenth; w.Divided = true; }
            }

            // D) Value clusters (gap split) -> primary cluster; chronological chain inside it catches single typos.
            decimal gap = hourVehicle ? 500m : 5000m;
            var byValue = live.OrderBy(w => w.Norm).ToList();
            var clusters = new List<List<OdoWork>>();
            foreach (var w in byValue)
            {
                if (clusters.Count == 0 || w.Norm - clusters[^1][^1].Norm > gap) clusters.Add(new List<OdoWork>());
                clusters[^1].Add(w);
            }

            List<OdoWork> primary = new();
            if (clusters.Count > 0)
            {
                decimal Dist(List<OdoWork> c) =>
                    masterOdo < c[0].Norm ? c[0].Norm - masterOdo : masterOdo > c[^1].Norm ? masterOdo - c[^1].Norm : 0m;
                var byMaster = masterOdo > 0m ? clusters.OrderBy(Dist).First() : null;
                primary = byMaster != null && Dist(byMaster) <= gap
                    ? byMaster
                    : clusters.OrderByDescending(c => c.Count).ThenByDescending(c => c.Max(w => OdoStamp(w.Log))).First();
            }
            var inPrimary = new HashSet<OdoWork>(primary);

            bool StepOk(OdoWork a, OdoWork b)
            {
                int days = Math.Max(0, b.Log.Date.DayNumber - a.Log.Date.DayNumber);
                decimal d = b.Norm - a.Norm;
                if (d < 0m) return false;
                if (d == 0m) return days == 0;
                decimal max = hourVehicle ? 60m + 14m * days : 1500m + 500m * days;
                return d <= max;
            }

            var chain = live.Where(inPrimary.Contains).ToList();   // already chronological
            var len = new int[chain.Count];
            var prev = new int[chain.Count];
            int bestEnd = -1, bestLen = 0;
            for (int i = 0; i < chain.Count; i++)
            {
                len[i] = 1; prev[i] = -1;
                for (int j = 0; j < i; j++)
                    if (len[j] + 1 > len[i] && StepOk(chain[j], chain[i])) { len[i] = len[j] + 1; prev[i] = j; }
                if (len[i] > bestLen) { bestLen = len[i]; bestEnd = i; }
            }
            var inChain = new HashSet<OdoWork>();
            for (int i = bestEnd; i >= 0; i = prev[i]) inChain.Add(chain[i]);

            var cand = new OdoCandidate
            {
                Plate = kv?.VehiclePlate ?? rows[0].OrigPlate,
                Key = key,
                Type = vType,
                Hour = hourVehicle,
                MasterOdo = masterOdo
            };
            if (bestLen >= 2 || live.Count == 1)
            {
                foreach (var w in chain.Where(inChain.Contains)) cand.Anchors.Add((OdoStamp(w.Log), w.Norm));
                cand.Anchors.Sort((x, y) => x.Stamp != y.Stamp ? x.Stamp.CompareTo(y.Stamp) : x.Odo.CompareTo(y.Odo));
            }
            candidates[key] = cand;

            // E) Outliers: other value clusters (card fuelled a different vehicle) and typos inside the primary cluster.
            foreach (var w in live)
            {
                if (w.PlateLocked) continue;
                if (!inPrimary.Contains(w)) { w.NeedsMatch = true; w.CanReject = true; }
                else if (bestLen >= 3 && !inChain.Contains(w)) { w.NeedsMatch = true; w.CanReject = true; }
            }
        }

        // Known vehicles without a timeline in the table still take part in matching via their master odometer.
        foreach (var entry in known)
        {
            if (candidates.ContainsKey(entry.Key)) continue;
            var t = entry.Value.VehicleDetail.Type;
            candidates[entry.Key] = new OdoCandidate
            {
                Plate = entry.Value.VehiclePlate,
                Key = entry.Key,
                Type = t,
                Hour = VehicleRecord.IsHourMeterType(t),
                MasterOdo = entry.Value.VehicleDetail.CurrentOdometer
            };
        }
        var candidateList = candidates.Values.Where(c => c.Anchors.Count > 0 || c.MasterOdo > 0m).ToList();

        // ---- Cross-fleet odometer trajectory matching (with ambiguity guard) -------------
        foreach (var w in works.Where(x => x.NeedsMatch && !x.Invalid && !x.PlateLocked).OrderBy(x => OdoStamp(x.Log)))
        {
            DateTime stamp = OdoStamp(w.Log);
            var hits = new List<(OdoCandidate C, decimal Odo, bool Div, int Tier, decimal Metric)>();

            foreach (var c in candidateList)
            {
                if (c.Key == w.Key || !FuelCompatible(w.Log, c)) continue;

                var readings = new List<(decimal Odo, bool Div)>();
                if (w.AlreadyScaled) readings.Add((w.Log.CurrentOdometer > 0m ? w.Log.CurrentOdometer : w.Norm, false));
                else
                {
                    readings.Add((w.Raw, false));
                    if (c.Hour) readings.Add((Math.Round(w.Raw / 10m, 1), true));
                }

                decimal bracketRate = c.Hour ? 20m : 800m;
                decimal boundRate = c.Hour ? 15m : 600m;
                decimal boundMax = c.Hour ? 150m : 3000m;
                decimal masterMax = c.Hour ? 100m : 2000m;

                (DateTime Stamp, decimal Odo)? before = null, after = null;
                foreach (var a in c.Anchors)
                {
                    if (a.Stamp <= stamp) before = a;
                    else { after = a; break; }
                }

                (OdoCandidate C, decimal Odo, bool Div, int Tier, decimal Metric)? bestForCandidate = null;
                foreach (var (odo, div) in readings)
                {
                    int tier = 0; decimal metric = 0m;
                    if (c.Anchors.Count > 0)
                    {
                        if (before != null && after != null)
                        {
                            if (odo >= before.Value.Odo && odo <= after.Value.Odo)
                            {
                                decimal d1 = odo - before.Value.Odo, d2 = after.Value.Odo - odo;
                                decimal dayB = Math.Max(1m, (decimal)(stamp - before.Value.Stamp).TotalDays);
                                decimal dayA = Math.Max(1m, (decimal)(after.Value.Stamp - stamp).TotalDays);
                                if (d1 / dayB <= bracketRate && d2 / dayA <= bracketRate) { tier = 1; metric = Math.Min(d1, d2); }
                            }
                        }
                        else if (before != null)
                        {
                            decimal d = odo - before.Value.Odo;
                            decimal days = (decimal)(stamp - before.Value.Stamp).TotalDays;
                            if (d > 0m && d <= boundMax && days <= 30m && d / Math.Max(1m, days) <= boundRate) { tier = 2; metric = d; }
                        }
                        else if (after != null)
                        {
                            decimal d = after.Value.Odo - odo;
                            decimal days = (decimal)(after.Value.Stamp - stamp).TotalDays;
                            if (d > 0m && d <= boundMax && days <= 30m && d / Math.Max(1m, days) <= boundRate) { tier = 2; metric = d; }
                        }
                    }
                    else if (c.MasterOdo > 0m)
                    {
                        decimal d = odo - c.MasterOdo;
                        if (d >= 0m && d <= masterMax) { tier = 3; metric = d; }
                    }
                    if (tier == 0) continue;

                    if (bestForCandidate == null || tier < bestForCandidate.Value.Tier
                        || (tier == bestForCandidate.Value.Tier && metric < bestForCandidate.Value.Metric))
                        bestForCandidate = (c, odo, div, tier, metric);
                }
                if (bestForCandidate != null) hits.Add(bestForCandidate.Value);
            }

            if (hits.Count == 0) continue;
            hits = hits.OrderBy(h => h.Tier).ThenBy(h => h.Metric).ToList();
            var best = hits[0];
            // One match, or a bracketed match that beats every looser match by more than 2x closeness.
            bool decisive = hits.Count == 1 ||
                (best.Tier == 1 && hits.Skip(1).All(h => h.Tier > 1 && h.Metric > 2m * best.Metric));

            if (!decisive)
            {
                // Never guess: keep the plate and ask for a human decision.
                AppendRemarkTag(w.Log,
                    $"[Ambiguous Odometer {RawText(w.Raw)}: matches {string.Join(", ", hits.Take(4).Select(h => h.C.Plate))} — verify vehicle]");
                w.Ambiguous = true;
                w.NeedsMatch = false;
                continue;
            }

            string orig = PlateTag(w.OrigPlate);
            w.Log.VehiclePlate = best.C.Plate;
            if (best.C.Type != VehicleType.None) w.Log.VehicleType = best.C.Type;
            w.Log.CurrentOdometer = best.Odo;
            if (best.Div)
            {
                AppendRemarkTag(w.Log, $"[{F1(best.Odo)} hrs (raw {F0(w.Raw)})]");
                hoursNormalized++;
            }
            AppendRemarkTag(w.Log, $"[Card {orig} -> Matched {PlateTag(best.C.Plate)} by Odometer]");
            w.Matched = true;
            w.PlateLocked = true;
            w.NeedsMatch = false;
            matchedCount++;
        }

        // ---- Apply remaining decisions ----------------------------------------------------
        foreach (var w in works)
        {
            if (w.Invalid || w.Matched) continue;

            if (w.NeedsMatch && w.CanReject)
            {
                Invalidate(w, $"raw {RawText(w.JunkRaw)}");
                continue;
            }
            if (w.Divided)
            {
                w.Log.CurrentOdometer = w.Norm;
                if (!ScaledHoursRx.IsMatch(w.Log.Remarks ?? string.Empty))
                {
                    AppendRemarkTag(w.Log, $"[{F1(w.Norm)} hrs (raw {F0(w.Raw)})]");
                    hoursNormalized++;
                }
            }
            else if (w.Log.CurrentOdometer <= 0m && w.Raw > 0m)
            {
                w.Log.CurrentOdometer = w.Raw;   // recovered reading that no longer needs rejecting
            }
        }

        // ---- Unrecognised plates / vehicle type ------------------------------------------
        foreach (var w in works)
        {
            var l = w.Log;
            if (IsKnownPlate(l.VehiclePlate))
            {
                // Populated on every row so PopulateTableSafe picks the right km / hrs display on first write.
                if (known.TryGetValue(VehiclePlateHelper.NormalizePlate(l.VehiclePlate), out var kvRec) &&
                    kvRec.VehicleDetail.Type != VehicleType.None)
                    l.VehicleType = kvRec.VehicleDetail.Type;
                continue;
            }

            // A blank plate would make TableMapper drop the row on the next extract; keep a visible placeholder.
            if (string.IsNullOrWhiteSpace(l.VehiclePlate)) l.VehiclePlate = "UNKNOWN";
            AppendRemarkTag(l, $"[Unrecognized Plate: {PlateTag(w.OrigPlate)} — please assign vehicle]");
        }

        // ---- Step D: chronological sort, ID repair, OdometerBefore chain ------------------
        var sorted = logs.OrderBy(l => l.Date).ThenBy(l => l.Time ?? string.Empty).ThenBy(l => l.OdometerID).ToList();
        logs.Clear();
        logs.AddRange(sorted);

        // IDs are database keys: only repair missing / duplicate ones, never renumber healthy rows.
        var seenIds = new HashSet<long>();
        long maxId = Math.Max(0, logs.Max(l => l.OdometerID));
        foreach (var l in logs)
        {
            if (l.OdometerID <= 0 || !seenIds.Add(l.OdometerID))
            {
                l.OdometerID = ++maxId;
                seenIds.Add(l.OdometerID);
            }
        }

        var lastValid = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        foreach (var l in logs)
        {
            string k = VehiclePlateHelper.NormalizePlate(l.VehiclePlate);
            decimal last = lastValid.TryGetValue(k, out var p) ? p : 0m;
            l.OdometerBefore = last;
            if (l.CurrentOdometer <= 0m) continue;   // lastValid carries forward unchanged

            if (last > 0m && l.CurrentOdometer < last)
            {
                // Regression that survived matching: keep it out of the forward chain.
                AppendRemarkTag(l, $"[Invalid Odometer: raw {RawText(l.CurrentOdometer)} < before {RawText(last)}]");
                l.CurrentOdometer = 0m;
                invalidZeroed++;
            }
            else
            {
                lastValid[k] = l.CurrentOdometer;
            }
        }

        int flagged = logs.Count(l => !IsKnownPlate(l.VehiclePlate) ||
                                      (l.Remarks ?? string.Empty).Contains("[Ambiguous Odometer", StringComparison.Ordinal));
        return new OdometerReconcileResult(logs.Count, matchedCount, hoursNormalized, invalidZeroed, flagged);
    }

    // =========================================================================
    // ODOMETERS TABLE COMMANDS & WRITE-BACK
    // =========================================================================

    /// <summary>Ribbon entry point: re-sorts and reconciles whatever is currently in OdometersTable.</summary>
    internal static OdometerReconcileResult SortAndHealOdometers()
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic wb = xlApp.ActiveWorkbook;
        if (wb == null) throw new InvalidOperationException("No active workbook found.");

        dynamic? wsOdo = SheetMapperHelper.ResolveSheetSafely((object)wb, OdometerAliases, "Odometers");
        if (wsOdo == null) throw new InvalidOperationException("Worksheet 'Odometers' was not found in active workbook.");

        // TableMapper skips rows without a plate; writing the list back would silently erase them.
        int platelessRows = CountRowsMissingPlate(wsOdo);
        if (platelessRows > 0)
            throw new InvalidOperationException(
                $"{platelessRows} row(s) in OdometersTable have no Vehicle Plate.\n" +
                "Enter a Vehicle Plate (or delete the rows) before sorting, otherwise they would be lost.");

        List<OdometerLog> logs = TableMapper.ExtractFromTable<OdometerLog>((object)wsOdo, "OdometersTable");
        if (logs.Count == 0) return new OdometerReconcileResult(0, 0, 0, 0, 0);

        List<FuelCardRecord> fuelCards = TableMapper.ExtractFromTable<FuelCardRecord>((object)wsOdo, "FuelCardsTable");
        var cardMap = new Dictionary<string, FuelCardRecord>(StringComparer.OrdinalIgnoreCase);
        foreach (var c in fuelCards)
        {
            string norm = SanitizeCardToken(c.FuelCardNumber);
            if (norm.Length > 0) cardMap[norm] = c;
        }

        dynamic? wsMaster = SheetMapperHelper.ResolveSheetSafely((object)wb, AllVehicleAliases, MasterSheetName);
        List<VehicleRecord> masters = wsMaster != null
            ? TableMapper.ExtractFromTable<VehicleRecord>((object)wsMaster, MasterTableName)
            : new List<VehicleRecord>();

        var result = ReconcileAndSortOdometerLogs(logs, masters, cardMap);
        WriteBackOdometers(xlApp, wsOdo, wsMaster, logs, masters);
        return result;
    }

    private static string SquashHeader(object? h) =>
        Regex.Replace((h?.ToString() ?? string.Empty).ToLowerInvariant(), "[^a-z0-9]", "");

    private static int FindColumn(dynamic table, string squashedName)
    {
        int cols = (int)table.ListColumns.Count;
        for (int c = 1; c <= cols; c++)
            if (SquashHeader(table.ListColumns[c].Name) == squashedName) return c;
        return 0;
    }

    private static int CountRowsMissingPlate(dynamic wsOdo)
    {
        dynamic? table = TableMapper.ResolveTable(wsOdo, "OdometersTable");
        dynamic? body = table?.DataBodyRange;
        if (table == null || body == null) return 0;

        int rows = (int)body.Rows.Count, cols = (int)table.ListColumns.Count;
        int plateCol = FindColumn(table, "vehicleplate");
        int dateCol = FindColumn(table, "date");
        int odoCol = FindColumn(table, "currentodometer");
        if (plateCol == 0) return 0;

        object[,] grid = TableMapper.GetRange2DValues(body, rows, cols);
        int count = 0;
        for (int r = 1; r <= rows; r++)
        {
            if (!string.IsNullOrWhiteSpace(grid[r, plateCol]?.ToString())) continue;
            bool hasData = (dateCol > 0 && !string.IsNullOrWhiteSpace(grid[r, dateCol]?.ToString()))
                        || (odoCol > 0 && !string.IsNullOrWhiteSpace(grid[r, odoCol]?.ToString()) && grid[r, odoCol]?.ToString() != "0");
            if (hasData) count++;
        }
        return count;
    }

    /// <summary>Writes the reconciled list into OdometersTable, then refreshes each vehicle's Current Odometer.</summary>
    private static void WriteBackOdometers(dynamic xlApp, dynamic wsOdo, dynamic? wsMaster, List<OdometerLog> logs, IList<VehicleRecord> masters)
    {
        SheetProtections.ExecuteWithBypass((object)wsOdo, new Action(() =>
        {
            TableMapper.PopulateTableSafe<OdometerLog>((object)wsOdo, logs, "OdometersTable");
            SheetProtections.ApplyStyleDrivenProtection(wsOdo);
            try { xlApp.Calculate(); } catch { }
        }));

        try
        {
            dynamic table = wsOdo.ListObjects["OdometersTable"];
            RuleEvaluator.EvaluateAndHealTable(wsOdo, table);
        }
        catch { }

        if (wsMaster == null) return;
        try { SyncMasterCurrentOdometers(xlApp, wsMaster, logs, masters); }
        catch (Exception ex) { AppLogger.Warn($"[Odometers] Master Current Odometer sync skipped: {ex.Message}"); }
    }

    /// <summary>Sets AllVehiclesTable[Current Odometer] to each vehicle's highest valid reading (one column write).</summary>
    private static void SyncMasterCurrentOdometers(dynamic xlApp, dynamic wsMaster, List<OdometerLog> logs, IList<VehicleRecord> masters)
    {
        var maxByPlate = logs
            .Where(l => l.CurrentOdometer > 0m && !string.IsNullOrWhiteSpace(l.VehiclePlate))
            .GroupBy(l => VehiclePlateHelper.NormalizePlate(l.VehiclePlate))
            .ToDictionary(g => g.Key, g => g.Max(x => x.CurrentOdometer), StringComparer.OrdinalIgnoreCase);
        if (maxByPlate.Count == 0) return;

        dynamic? table = TableMapper.ResolveTable(wsMaster, MasterTableName);
        dynamic? body = table?.DataBodyRange;
        if (table == null || body == null) return;

        int plateCol = FindColumn(table, "vehicleplate");
        int odoCol = FindColumn(table, "currentodometer");
        if (plateCol == 0 || odoCol == 0) return;

        int rows = (int)body.Rows.Count;
        dynamic plateRange = table.ListColumns[plateCol].DataBodyRange;
        dynamic odoRange = table.ListColumns[odoCol].DataBodyRange;
        if ((odoRange.Cells[1, 1].Formula?.ToString() ?? string.Empty).StartsWith("=")) return;

        object[,] plates = TableMapper.GetRange2DValues(plateRange, rows, 1);
        object[,] current = TableMapper.GetRange2DValues(odoRange, rows, 1);

        var output = new object[rows, 1];
        bool changed = false;
        for (int r = 0; r < rows; r++)
        {
            object? existing = current[r + 1, 1];
            object? value = existing;
            string k = VehiclePlateHelper.NormalizePlate(plates[r + 1, 1]?.ToString());
            if (maxByPlate.TryGetValue(k, out decimal max))
            {
                double newVal = (double)max;
                if (existing is not double old || Math.Abs(old - newVal) > 0.0001) changed = true;
                value = newVal;
            }
            output[r, 0] = value!;
        }
        var typeByPlate = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        foreach (var m in masters)
            typeByPlate[VehiclePlateHelper.NormalizePlate(m.VehiclePlate)] = VehicleRecord.IsHourMeterType(m.VehicleDetail.Type);
        var isHours = new List<bool>(rows);
        for (int r = 0; r < rows; r++)
        {
            string k = VehiclePlateHelper.NormalizePlate(plates[r + 1, 1]?.ToString());
            isHours.Add(typeByPlate.TryGetValue(k, out bool h) && h);
        }

        bool origEvents = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try
        {
            SheetProtections.ExecuteWithBypass((object)wsMaster, new Action(() =>
            {
                if (changed) odoRange.Value2 = output;
                ApplyUnitFormatRuns(wsMaster, odoRange, isHours);
            }));
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    /// <summary>Applies km / hrs NumberFormat per contiguous run of rows sharing a unit (one COM call per run).</summary>
    private static void ApplyUnitFormatRuns(dynamic ws, dynamic columnRange, IList<bool> isHours)
    {
        int start = 0;
        while (start < isHours.Count)
        {
            int end = start;
            while (end + 1 < isHours.Count && isHours[end + 1] == isHours[start]) end++;
            try
            {
                dynamic run = ws.Range[columnRange.Cells[start + 1, 1], columnRange.Cells[end + 1, 1]];
                run.NumberFormat = isHours[start] ? "#,##0.# \"hrs\"" : "#,##0 \"km\"";
            }
            catch { }
            start = end + 1;
        }
    }

    // =========================================================================
    // CARD & STAFF RESOLUTION HELPERS
    // =========================================================================

    private static string SanitizeCardToken(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        // Strip quotes, spaces, and hyphens, keeping digits and '*'
        return Regex.Replace(raw.Trim(), @"['\s\-]+", "");
    }

    private static string NormalizeName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return string.Empty;
        return Regex.Replace(name.Trim(), @"\s+", " ");
    }

    /// <summary>
    /// Matches a raw or '*' masked card number against FuelCardsTable.
    /// Supports single or multiple '*' blocks (e.g. 7077841******000208) and returns the canonical card number.
    /// </summary>
    private static string ResolveFullCardNumber(
        string rawCard,
        Dictionary<string, FuelCardRecord> cardMap,
        out FuelCardRecord? matchedRecord)
    {
        matchedRecord = null;
        string clean = SanitizeCardToken(rawCard);
        if (string.IsNullOrWhiteSpace(clean)) return string.Empty;

        // 1. Direct unmasked match
        if (cardMap.TryGetValue(clean, out var exact))
        {
            matchedRecord = exact;
            return exact.FuelCardNumber.Trim().Replace("'", "");
        }

        // 2. Wildcard '*' match against FuelCardsTable
        if (clean.Contains('*'))
        {
            // Build a regex where any sequence of '*' matches 1 or more characters
            string pattern = "^" + string.Join(".+", clean.Split('*', StringSplitOptions.None).Select(Regex.Escape)) + "$";
            var wildcardRx = new Regex(pattern, RegexOptions.IgnoreCase);

            // Also extract simple prefix/suffix as a fallback if lengths differ slightly
            string prefix = clean.Substring(0, clean.IndexOf('*'));
            string suffix = clean.Substring(clean.LastIndexOf('*') + 1);

            foreach (var (normFullNum, cardRec) in cardMap)
            {
                if (wildcardRx.IsMatch(normFullNum) ||
                    (!string.IsNullOrEmpty(prefix) && !string.IsNullOrEmpty(suffix) &&
                     normFullNum.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                     normFullNum.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)))
                {
                    matchedRecord = cardRec;
                    return cardRec.FuelCardNumber.Trim().Replace("'", "");
                }
            }
        }

        return clean;
    }

    /// <summary>
    /// Resolves a driver name from the statement to a StaffID using multi-tier matching:
    /// 1. Exact match on StaffShortName or StaffName
    /// 2. Whole-word / token match (e.g. "SYUKUR" matches "MOHD SYUKUR")
    /// 3. Substring fallback
    /// </summary>
    private static long ResolveStaffId(
        string driverName,
        Dictionary<string, long> staffExactMap,
        List<StaffRecord> staffList)
    {
        string clean = NormalizeName(driverName);
        if (string.IsNullOrWhiteSpace(clean)) return 0;

        // Tier 1: Exact match against ShortName or FullName
        if (staffExactMap.TryGetValue(clean, out long exactId))
            return exactId;

        var driverTokens = clean.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        // Tier 2: Whole-word token match (prioritize StaffShortName first, then StaffName)
        long bestStaffId = 0;
        int bestMatchedWords = 0;

        foreach (var s in staffList)
        {
            if (s.StaffID <= 0) continue;

            string sn = NormalizeName(s.StaffDetails?.StaffShortName);
            string fn = NormalizeName(s.StaffDetails?.StaffName);

            foreach (var candidate in new[] { sn, fn })
            {
                if (string.IsNullOrEmpty(candidate)) continue;
                var candidateTokens = candidate.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                // Check if all tokens in driverName exist as whole words in candidate (or vice versa)
                bool driverInCandidate = driverTokens.All(dt =>
                    candidateTokens.Any(ct => ct.Equals(dt, StringComparison.OrdinalIgnoreCase)));
                bool candidateInDriver = candidateTokens.All(ct =>
                    driverTokens.Any(dt => dt.Equals(ct, StringComparison.OrdinalIgnoreCase)));

                if (driverInCandidate || candidateInDriver)
                {
                    int overlap = driverTokens.Count(dt =>
                        candidateTokens.Any(ct => ct.Equals(dt, StringComparison.OrdinalIgnoreCase)));
                    if (overlap > bestMatchedWords)
                    {
                        bestMatchedWords = overlap;
                        bestStaffId = s.StaffID;
                    }
                }
            }
        }

        if (bestStaffId > 0) return bestStaffId;

        // Tier 3: Substring containment fallback (only for names >= 3 chars to avoid false positives)
        if (clean.Length >= 3)
        {
            foreach (var (name, id) in staffExactMap)
            {
                if (name.Length >= 3 &&
                    (name.Contains(clean, StringComparison.OrdinalIgnoreCase) ||
                     clean.Contains(name, StringComparison.OrdinalIgnoreCase)))
                {
                    return id;
                }
            }
        }

        return 0;
    }
}

// =============================================================================
// INTERNAL STATEMENT PARSER (ROTATION-AWARE LETTER & RIBBON ENGINE)
// =============================================================================
internal sealed class RawFuelItem
{
    public DateOnly Date { get; set; }
    public string Time { get; set; } = "00:00";
    public string VehiclePlate { get; set; } = string.Empty;
    public string VehicleType { get; set; } = string.Empty;
    public long Odometer { get; set; }
    public string CardNumber { get; set; } = string.Empty;
    public long StaffID { get; set; }
    public string StaffName { get; set; } = string.Empty;
    public string StationName { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public string ReceiptNumber { get; set; } = string.Empty;
    public string FuelType { get; set; } = string.Empty;
    public decimal PriceAtFueling { get; set; }
    public decimal FuelAmount { get; set; }
}

internal static class FuelStatementParser
{
    private static readonly Regex PlateRegex = new(@"\b([A-Z]{1,3})\s*(\d{1,4})\s*([A-Z]?)\b", RegexOptions.Compiled);

    public static List<RawFuelItem> Parse(string filePath, List<string> knownDriverNames, out string providerName)
    {
        providerName = "Unknown Portal";
        var list = new List<RawFuelItem>();
        if (!File.Exists(filePath))
        {
            AppLogger.Error($"[FuelStatementParser] FILE NOT FOUND: {filePath}");
            return list;
        }

        // Order known driver names by length descending so multi-word names match before single-word subsets
        var sortedKnownNames = (knownDriverNames ?? new List<string>())
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(n => Regex.Replace(n.Trim(), @"\s+", " "))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(n => n.Length)
            .ToList();

        AppLogger.Info($"[FuelStatementParser] === STARTING PARSE FOR: {Path.GetFileName(filePath)} ===");
        string ext = Path.GetExtension(filePath).ToLowerInvariant();

        if (ext == ".pdf")
        {
            try
            {
                using var doc = PdfDocument.Open(filePath);
                var sb = new StringBuilder();
                foreach (var p in doc.GetPages()) sb.Append(p.Text);
                string full = sb.ToString();

                if (full.Contains("PETRONAS", StringComparison.OrdinalIgnoreCase) ||
                    full.Contains("smartPAY", StringComparison.OrdinalIgnoreCase))
                {
                    providerName = "PETRONAS SmartPay";
                    AppLogger.Info("[FuelStatementParser] Detected Format: PETRONAS SmartPay PDF");
                    list = ParsePetronasPdf(doc, sortedKnownNames);
                }
                else if (full.Contains("Shell Malaysia Trading", StringComparison.OrdinalIgnoreCase) ||
                         full.Contains("fleethub.shell.com", StringComparison.OrdinalIgnoreCase))
                {
                    providerName = "Shell Fleet Hub";
                    AppLogger.Info("[FuelStatementParser] Detected Format: Shell Fleet Hub PDF");
                    list = ParseShellPdf(doc);
                }
                else
                {
                    AppLogger.Warn("[FuelStatementParser] Failed to detect PETRONAS or SHELL keywords in PDF text.");
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error($"[FuelStatementParser] PDF Read Exception: {ex.Message}", ex);
            }
        }
        else if (ext == ".xlsx" || ext == ".xls" || ext == ".csv")
        {
            providerName = "Spreadsheet Portal Export";
            AppLogger.Info("[FuelStatementParser] Detected Format: Generic Spreadsheet Export");
            list = ParseSpreadsheet(filePath);
        }

        AppLogger.Info($"[FuelStatementParser] Total Records Successfully Parsed: {list.Count}");
        return list;
    }

    private static List<string> ExtractLinesFromLetters(Page page, double yTolerance = 3.2)
    {
        var letters = page.Letters;
        if (letters == null || letters.Count == 0) return new List<string>();

        double sumAbsDx = 0, sumAbsDy = 0;
        double sumSignedDx = 0, sumSignedDy = 0;
        int sampleLimit = Math.Min(letters.Count - 1, 200);

        for (int i = 0; i < sampleLimit; i++)
        {
            double dx = letters[i + 1].StartBaseLine.X - letters[i].StartBaseLine.X;
            double dy = letters[i + 1].StartBaseLine.Y - letters[i].StartBaseLine.Y;
            double dist = Math.Abs(dx) + Math.Abs(dy);

            if (dist > 0.4 && dist < 15.0)
            {
                sumAbsDx += Math.Abs(dx);
                sumAbsDy += Math.Abs(dy);
                sumSignedDx += dx;
                sumSignedDy += dy;
            }
        }

        bool horizontalText = sumAbsDx >= sumAbsDy;
        double colSign = horizontalText
            ? (sumSignedDx >= 0 ? 1.0 : -1.0)
            : (sumSignedDy >= 0 ? 1.0 : -1.0);

        double GetColCoord(PdfPoint pt) => horizontalText ? (pt.X * colSign) : (pt.Y * colSign);
        double GetRawRow(Letter l) => horizontalText ? l.StartBaseLine.Y : l.StartBaseLine.X;

        int edgeCount = Math.Min(25, letters.Count);
        double firstRowAvg = letters.Take(edgeCount).Average(GetRawRow);
        double lastRowAvg = letters.Skip(letters.Count - edgeCount).Average(GetRawRow);
        double rowSign = firstRowAvg <= lastRowAvg ? 1.0 : -1.0;

        double GetRowCoord(Letter l) => GetRawRow(l) * rowSign;

        var sortedByRow = letters.OrderBy(GetRowCoord).ToList();
        var clusters = new List<List<Letter>>();

        foreach (var lt in sortedByRow)
        {
            if (string.IsNullOrEmpty(lt.Value)) continue;
            double r = GetRowCoord(lt);
            var existing = clusters.FirstOrDefault(c => Math.Abs(GetRowCoord(c[0]) - r) <= yTolerance);
            if (existing != null)
            {
                existing.Add(lt);
            }
            else
            {
                clusters.Add(new List<Letter> { lt });
            }
        }

        var resultLines = new List<string>(clusters.Count);
        foreach (var cluster in clusters)
        {
            var ordered = cluster.OrderBy(l => GetColCoord(l.StartBaseLine)).ToList();
            var sb = new StringBuilder();

            for (int i = 0; i < ordered.Count; i++)
            {
                var curr = ordered[i];
                if (i > 0)
                {
                    var prev = ordered[i - 1];
                    double prevStart = GetColCoord(prev.StartBaseLine);
                    double prevEnd = GetColCoord(prev.EndBaseLine);
                    double currStart = GetColCoord(curr.StartBaseLine);

                    double charAdvance = Math.Abs(prevEnd - prevStart);
                    double gap;
                    if (charAdvance > 0.2 && charAdvance < 25.0)
                    {
                        gap = currStart - Math.Max(prevStart, prevEnd);
                    }
                    else
                    {
                        double fallbackWidth = Math.Min(prev.GlyphRectangle.Width, prev.GlyphRectangle.Height);
                        if (fallbackWidth < 1.0) fallbackWidth = 4.2;
                        gap = (currStart - prevStart) - fallbackWidth;
                    }

                    if (gap > 1.2 && !prev.Value.EndsWith(" ") && !curr.Value.StartsWith(" "))
                    {
                        sb.Append(' ');
                    }
                }
                sb.Append(curr.Value);
            }

            string cleanLine = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
            if (!string.IsNullOrWhiteSpace(cleanLine))
            {
                resultLines.Add(cleanLine);
            }
        }

        return resultLines;
    }

    private static string GetContinuousTextRibbon(Page page, double yTolerance = 4.0)
    {
        var words = page.GetWords().ToList();
        if (words.Count == 0) return string.Empty;

        var sortedWords = words.OrderByDescending(w => w.BoundingBox.Bottom).ThenBy(w => w.BoundingBox.Left).ToList();
        var lines = new List<string>();
        var currentLine = new List<Word>();
        double currentY = sortedWords[0].BoundingBox.Bottom;

        foreach (var word in sortedWords)
        {
            if (Math.Abs(word.BoundingBox.Bottom - currentY) > yTolerance)
            {
                if (currentLine.Count > 0)
                {
                    lines.Add(string.Join(" ", currentLine.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text)));
                    currentLine.Clear();
                }
                currentY = word.BoundingBox.Bottom;
            }
            currentLine.Add(word);
        }
        if (currentLine.Count > 0)
        {
            lines.Add(string.Join(" ", currentLine.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text)));
        }
        return string.Join(" ", lines);
    }

    // =========================================================================
    // 1. PETRONAS SMARTPAY PARSER
    // =========================================================================
    private static List<RawFuelItem> ParsePetronasPdf(PdfDocument doc, List<string> sortedKnownNames)
    {
        var items = new List<RawFuelItem>();

        // Supports both full digits and '*' masked cards if Petronas ever masks digits
        var cardRx = new Regex(@"Nombor\s*Kad\s*[:\s]*([\d\*]{12,22})", RegexOptions.IgnoreCase);
        var plateRx = new Regex(@"Pendaftaran\s*Kenderaan\s*[:\s]*([A-Z0-9]+)", RegexOptions.IgnoreCase);

        var fullRowRx = new Regex(
            @"(\d{2}/\d{2}/\d{4})\s*(\d{2}:\d{2}(?::\d{2})?)\s*(\d{8,15})\s*(.+?)\s*(00\d{8})\s*((?:Diesel|PRIMAX).*?)\s+([\d\.,]+\s+[\d\.,]+\s+[\-\d\.,]+\s+[\-\d\.,]+\s+[\-\d\.,]+)",
            RegexOptions.IgnoreCase);

        string carryCard = string.Empty;
        string carryPlate = string.Empty;

        foreach (var page in doc.GetPages())
        {
            string rawText = page.Text ?? string.Empty;
            var lines = ExtractLinesFromLetters(page, yTolerance: 3.2);
            string joinedLines = string.Join(" | ", lines);

            if (!rawText.Contains("LAPORAN ANALISA KENDERAAN", StringComparison.OrdinalIgnoreCase) &&
                !rawText.Contains("ODOMETER", StringComparison.OrdinalIgnoreCase) &&
                !joinedLines.Contains("ODOMETER", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var cm = cardRx.Match(joinedLines);
            if (!cm.Success) cm = cardRx.Match(rawText);
            if (cm.Success) carryCard = cm.Groups[1].Value.Trim();

            var pm = plateRx.Match(joinedLines);
            if (!pm.Success) pm = plateRx.Match(rawText);
            if (pm.Success)
            {
                carryPlate = Regex.Replace(pm.Groups[1].Value.Trim(), @"(Nombor|Kod|TARIKH|Mukasurat).*$", "", RegexOptions.IgnoreCase).Trim();
            }

            foreach (var line in lines)
            {
                var m = fullRowRx.Match(line);
                if (!m.Success) continue;

                if (!DateOnly.TryParseExact(m.Groups[1].Value, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var txDate))
                    continue;

                string timeStr = m.Groups[2].Value;
                if (timeStr.Length >= 5) timeStr = timeStr.Substring(0, 5);

                string receipt = m.Groups[3].Value;
                string middle = m.Groups[4].Value.Trim();
                long.TryParse(m.Groups[5].Value, out long odo);

                string rawFuelType = Regex.Replace(m.Groups[6].Value, @"\s+", " ").Trim();

                var numTokens = m.Groups[7].Value.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                string netStr = numTokens.Last().TrimEnd('.').Replace(",", "");
                decimal.TryParse(netStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal netAmt);
                decimal.TryParse(numTokens[0].Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal fuelQty);
                decimal unitPrice = 0m;
                if (numTokens.Length > 1)
                    decimal.TryParse(numTokens[1].Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out unitPrice);
                if (unitPrice <= 0m && fuelQty > 0m && netAmt > 0m)
                    unitPrice = Math.Round(netAmt / fuelQty, 2);

                ExtractPetronasStationAndDriver(middle, sortedKnownNames, out string station, out string driver);

                string finalPlate = carryPlate;
                if ((finalPlate.Equals("SPARE", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(finalPlate)) &&
                    PlateRegex.IsMatch(driver))
                {
                    finalPlate = driver;
                }

                AppLogger.Info($"[PETRONAS MATCH] P{page.Number} Date:{txDate} Time:{timeStr} Plate:{finalPlate} Driver:{driver} Station:{station} Odo:{odo} Amt:{netAmt}");

                items.Add(new RawFuelItem
                {
                    Date = txDate,
                    Time = timeStr,
                    VehiclePlate = finalPlate,
                    CardNumber = carryCard,
                    StaffName = driver,
                    StationName = station,
                    Odometer = odo,
                    Amount = netAmt,
                    ReceiptNumber = receipt,
                    FuelType = rawFuelType,
                    PriceAtFueling = unitPrice,
                    FuelAmount = fuelQty
                });
            }
        }

        return items;
    }

    private static void ExtractPetronasStationAndDriver(
        string middle,
        List<string> sortedKnownNames,
        out string station,
        out string driver)
    {
        station = "PETRONAS STATION";
        driver = string.Empty;
        if (string.IsNullOrWhiteSpace(middle)) return;

        string clean = Regex.Replace(middle, @"\s+", " ").Trim();
        clean = Regex.Replace(clean, @"^(PSS)(?=[A-Z0-9])", "$1 ");

        // 1. Check if the driver column holds a vehicle plate (e.g. SPARE cards)
        var plateMatch = PlateRegex.Match(clean);
        if (plateMatch.Success && clean.EndsWith(plateMatch.Value, StringComparison.OrdinalIgnoreCase))
        {
            driver = plateMatch.Value.Trim();
            station = clean.Substring(0, clean.Length - driver.Length).Trim();
            return;
        }

        // 2. Match against known StaffShortName / StaffName from Staff Directory (longest match first)
        foreach (var known in sortedKnownNames)
        {
            if (clean.Length > known.Length &&
                clean.EndsWith(" " + known, StringComparison.OrdinalIgnoreCase))
            {
                driver = clean.Substring(clean.Length - known.Length).Trim();
                station = clean.Substring(0, clean.Length - known.Length).Trim();
                return;
            }
        }

        // 3. Fallback heuristic for unlisted multi-word Malaysian names
        var tokens = clean.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 0) return;

        int driverWords = 1;
        if (tokens.Length >= 3)
        {
            string tPrev = tokens[tokens.Length - 2];
            if (tPrev.Equals("MOHD", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("MUHAMMAD", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("MUHD", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("ABDUL", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("ABD", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("AHMAD", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("SYED", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("WAN", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("NIK", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("LEE", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("TAN", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("LIM", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("WONG", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("NG", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("YAP", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("CHEN", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("CHONG", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("CHAN", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("ONG", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("KUM", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("KENG", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("KIAN", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("SAIFUL", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("FEKRI", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("SRI", StringComparison.OrdinalIgnoreCase) ||
                tPrev.Equals("HANIF", StringComparison.OrdinalIgnoreCase))
            {
                driverWords = 2;
                if (tokens.Length >= 4)
                {
                    string tPrev2 = tokens[tokens.Length - 3];
                    if (tPrev2.Equals("CHEN", StringComparison.OrdinalIgnoreCase) ||
                        tPrev2.Equals("LEE", StringComparison.OrdinalIgnoreCase) ||
                        tPrev2.Equals("TAN", StringComparison.OrdinalIgnoreCase) ||
                        tPrev2.Equals("LIM", StringComparison.OrdinalIgnoreCase) ||
                        tPrev2.Equals("WONG", StringComparison.OrdinalIgnoreCase) ||
                        tPrev2.Equals("YAP", StringComparison.OrdinalIgnoreCase) ||
                        tPrev2.Equals("CHONG", StringComparison.OrdinalIgnoreCase))
                    {
                        driverWords = 3;
                    }
                }
            }
        }

        driver = string.Join(" ", tokens.Skip(tokens.Length - driverWords));
        station = string.Join(" ", tokens.Take(tokens.Length - driverWords));
    }

    // =========================================================================
    // 2. SHELL FLEET HUB PARSER
    // =========================================================================
    private static List<RawFuelItem> ParseShellPdf(PdfDocument doc)
    {
        var items = new List<RawFuelItem>();

        // Flexible masked/unmasked Shell card regex (handles any number of '*' or spaces)
        var shellCardRx = new Regex(@"(70\d{2,6}[\*\s]*\*{2,10}[\*\s]*\d{3,6})\s+([A-Z\s]+?)\s+([A-Z]{1,3}\s*\d{1,4}\s*[A-Z]?)\b", RegexOptions.IgnoreCase);

        var shellTxRx = new Regex(
            @"(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})\s+(.+?)\s+(\d{5,8})\s+(\d+)\s+(\d{2})\s+([\d\.,]+)\s*Lit\s+([\d\.,]+)\s+([\-\d\.,]+)(?:\s+(-\d[\d\.,]*))?(?:\s+(-\d[\d\.,]*))?",
            RegexOptions.IgnoreCase);

        string carryCard = string.Empty;
        string carryDriver = string.Empty;
        string carryPlate = string.Empty;

        foreach (var page in doc.GetPages())
        {
            string pageText = GetContinuousTextRibbon(page);
            if (!pageText.Contains("TRANSACTION DETAILS") && !pageText.Contains("BUTIRAN TRANSAKSI"))
                continue;

            var cardMatches = shellCardRx.Matches(pageText);
            var txMatches = shellTxRx.Matches(pageText);

            foreach (Match tm in txMatches)
            {
                if (!DateOnly.TryParseExact(tm.Groups[1].Value, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var txDate))
                    continue;

                string time = tm.Groups[2].Value;
                string receipt = tm.Groups[4].Value;
                long.TryParse(tm.Groups[5].Value, out long odo);

                decimal grossAmt = 0m;
                decimal subsidyAmt = 0m;

                for (int i = 9; i <= 11; i++)
                {
                    if (tm.Groups[i].Success && !string.IsNullOrWhiteSpace(tm.Groups[i].Value))
                    {
                        string rawNum = tm.Groups[i].Value.Replace(",", "");
                        int dotIdx = rawNum.IndexOf('.');
                        int decimals = dotIdx >= 0 ? rawNum.Length - dotIdx - 1 : 0;
                        if (decimals != 2) continue;

                        if (decimal.TryParse(rawNum, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal val))
                        {
                            if (val >= 0 && grossAmt == 0m) grossAmt = val;
                            else if (val < 0 && subsidyAmt == 0m) subsidyAmt = val;
                        }
                    }
                }

                decimal.TryParse(tm.Groups[7].Value.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal fuelQty);
                decimal.TryParse(tm.Groups[8].Value.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal unitPrice);
                if (unitPrice <= 0m && fuelQty > 0m && grossAmt > 0m)
                    unitPrice = Math.Round(grossAmt / fuelQty, 2);

                string shellCode = tm.Groups[6].Value;
                string shellDesc = tm.Groups[3].Value;
                string shellFuelType;
                if (shellDesc.Contains("V-Power", StringComparison.OrdinalIgnoreCase) || shellDesc.Contains("97")) shellFuelType = "RON 97";
                else if (shellDesc.Contains("95")) shellFuelType = "RON 95";
                else if (shellDesc.Contains("Diesel", StringComparison.OrdinalIgnoreCase) || subsidyAmt < 0m ||
                         shellCode is "20" or "23" or "31") shellFuelType = "Diesel";
                else shellFuelType = "Shell Code " + shellCode;

                decimal netAmount = grossAmt + subsidyAmt;
                if (netAmount <= 0m && grossAmt > 0m) netAmount = grossAmt;

                string station = Regex.Replace(tm.Groups[3].Value, @"^\s*\d+\s*", "").Trim();
                station = CleanStation(station);

                Match? nearestCard = null;
                foreach (Match cm in cardMatches)
                {
                    if (cm.Index < tm.Index) nearestCard = cm;
                }

                if (nearestCard != null)
                {
                    carryCard = Regex.Replace(nearestCard.Groups[1].Value, @"\s+", "");
                    carryDriver = nearestCard.Groups[2].Value.Trim();
                    carryPlate = nearestCard.Groups[3].Value.Trim();
                }

                AppLogger.Info($"[SHELL MATCH] Date:{txDate} Card:{carryCard} Plate:{carryPlate} Driver:{carryDriver} Odo:{odo} Amt:{netAmount}");

                items.Add(new RawFuelItem
                {
                    Date = txDate,
                    Time = time,
                    VehiclePlate = carryPlate,
                    CardNumber = carryCard,
                    StaffName = carryDriver,
                    StationName = station,
                    Odometer = odo,
                    Amount = netAmount,
                    ReceiptNumber = receipt,
                    FuelType = shellFuelType,
                    PriceAtFueling = unitPrice,
                    FuelAmount = fuelQty
                });
            }

            if (cardMatches.Count > 0)
            {
                var lastCard = cardMatches[cardMatches.Count - 1];
                carryCard = Regex.Replace(lastCard.Groups[1].Value, @"\s+", "");
                carryDriver = lastCard.Groups[2].Value.Trim();
                carryPlate = lastCard.Groups[3].Value.Trim();
            }
        }
        return items;
    }

    // =========================================================================
    // 3. SPREADSHEET FALLBACK (.xlsx / .xls / .csv)
    // =========================================================================
    private static List<RawFuelItem> ParseSpreadsheet(string path)
    {
        var items = new List<RawFuelItem>();
        var lines = File.ReadAllLines(path);
        if (lines.Length == 0) return items;

        char delim = lines[0].Contains('\t') ? '\t' : ',';
        int headerRow = -1;
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < Math.Min(15, lines.Length); i++)
        {
            var cols = lines[i].Split(delim);
            for (int c = 0; c < cols.Length; c++)
            {
                string h = cols[c].Trim().ToLowerInvariant();
                if (h.Contains("date") || h.Contains("card") || h.Contains("amount"))
                {
                    headerRow = i;
                    break;
                }
            }
            if (headerRow >= 0) break;
        }

        if (headerRow < 0) return items;

        var headerCols = lines[headerRow].Split(delim);
        for (int c = 0; c < headerCols.Length; c++)
            map[headerCols[c].Trim().Replace("\"", "").ToLowerInvariant()] = c;

        for (int i = headerRow + 1; i < lines.Length; i++)
        {
            string line = lines[i].Trim();
            if (string.IsNullOrWhiteSpace(line) || line.Contains("TOTAL")) continue;

            var p = line.Split(delim);
            string dStr = GetVal(p, map, "date");
            if (string.IsNullOrWhiteSpace(dStr)) continue;

            if (DateTime.TryParse(dStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt) ||
                DateTime.TryParse(dStr, new CultureInfo("en-GB"), DateTimeStyles.None, out dt))
            {
                decimal.TryParse(GetAmountVal(p, map).Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amt);
                decimal.TryParse(GetValAny(p, map, "price", "unit price").Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal price);
                decimal.TryParse(GetValAny(p, map, "fuel amount", "lit", "qty", "quantity", "volume").Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal fuelQty);
                string timeVal = GetValAny(p, map, "time");
                if (!string.IsNullOrWhiteSpace(timeVal) && TimeOnly.TryParse(timeVal, CultureInfo.InvariantCulture, DateTimeStyles.None, out var tParsed))
                    timeVal = tParsed.ToString("HH:mm", CultureInfo.InvariantCulture);
                else
                    timeVal = dt.ToString("HH:mm");
                long.TryParse(GetVal(p, map, "odometer"), out long odo);

                items.Add(new RawFuelItem
                {
                    Date = DateOnly.FromDateTime(dt),
                    Time = timeVal,
                    CardNumber = GetVal(p, map, "card").Replace("'", ""),
                    VehiclePlate = GetVal(p, map, "vehicle"),
                    StaffName = GetVal(p, map, "driver"),
                    StationName = GetVal(p, map, "station"),
                    Amount = amt,
                    Odometer = odo,
                    FuelType = GetValAny(p, map, "fuel type", "product"),
                    PriceAtFueling = price,
                    FuelAmount = fuelQty,
                    ReceiptNumber = GetValAny(p, map, "receipt", "reciept")
                });
            }
        }
        return items;
    }

    private static string CleanStation(string raw)
    {
        string s = Regex.Replace(raw, @"\s+", " ").Trim();
        int slash = s.IndexOf('/');
        return slash > 0 ? s.Substring(0, slash).Trim() : s;
    }

    private static string GetValAny(string[] p, Dictionary<string, int> map, params string[] keys)
    {
        foreach (var key in keys)
        {
            string v = GetVal(p, map, key);
            if (!string.IsNullOrWhiteSpace(v)) return v;
        }
        return string.Empty;
    }

    // "amount" must not pick up the "Fuel Amount" (litres) column.
    private static string GetAmountVal(string[] p, Dictionary<string, int> map)
    {
        foreach (var (k, idx) in map)
        {
            if (k.Contains("amount") && !k.Contains("fuel") && idx < p.Length)
                return p[idx].Trim().Replace("\"", "");
        }
        return string.Empty;
    }

    private static string GetVal(string[] p, Dictionary<string, int> map, string key)
    {
        foreach (var (k, idx) in map)
        {
            if (k.Contains(key) && idx < p.Length)
                return p[idx].Trim().Replace("\"", "");
        }
        return string.Empty;
    }
}