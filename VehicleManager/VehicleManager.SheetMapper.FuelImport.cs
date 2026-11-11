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
    private static string BuildFuelSignature(OdometerLog log)
    {
        string id = string.IsNullOrWhiteSpace(log.VehiclePlate)
            ? (log.FuelCardNumber ?? string.Empty).Trim()
            : VehiclePlateHelper.NormalizePlate(log.VehiclePlate);
        return $"{log.Date:yyyyMMdd}_{(log.Time ?? string.Empty).Trim()}_{id}_{log.Amount:F2}_{(log.ReceiptNumber ?? string.Empty).Trim()}_{log.CurrentOdometer}_{(log.StationName ?? string.Empty).Trim()}_{(log.FuelType ?? string.Empty).Trim()}_{log.PriceAtFueling:F2}_{log.FuelAmount:F2}";
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
                CurrentOdometer = raw.Odometer > 0 ? (int)raw.Odometer : 0,
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
            SetProp(log, "FuelTransactionID", maxTxId);
            SetProp(log, "TransactionID", maxTxId);
            totalAmount += GetProp<decimal>(log, "Amount", "TransactionAmount");
        }

        if (cleanTransactions.Count > 0)
        {
            var combined = new List<OdometerLog>(existingLogs.Count + cleanTransactions.Count);
            combined.AddRange(existingLogs);
            combined.AddRange(cleanTransactions);

            combined = combined.OrderBy(x => GetProp<DateOnly>(x, "Date"))
                               .ThenBy(x => x.Time ?? string.Empty)
                               .ThenBy(x => x.OdometerID)
                               .ToList();

            SheetProtections.ExecuteWithBypass((object)wsOdo, new Action(() =>
            {
                TableMapper.PopulateTableSafe<OdometerLog>((object)wsOdo, combined, "OdometersTable");
                SheetProtections.ApplyStyleDrivenProtection(wsOdo);
                try { xlApp.Calculate(); } catch { }
            }));

            try
            {
                dynamic table = wsOdo.ListObjects["OdometersTable"];
                RuleEvaluator.EvaluateAndHealTable(wsOdo, table);
            }
            catch { }
        }

        return (cleanTransactions.Count, skippedCount, totalAmount, providerName);
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