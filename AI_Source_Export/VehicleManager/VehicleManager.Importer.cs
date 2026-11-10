namespace SKSB_App.VehicleManager.Importer;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

public enum FuelProvider
{
    Unknown,
    Petronas,
    Shell
}

public sealed class FuelTransactionRecord
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
    public string ProductDescription { get; set; } = string.Empty;
    public decimal Litres { get; set; }
    public FuelProvider Provider { get; set; }
}

public sealed class FuelImportResult
{
    public FuelProvider Provider { get; set; } = FuelProvider.Unknown;
    public string SourceFileName { get; set; } = string.Empty;
    public List<FuelTransactionRecord> NewTransactions { get; } = new();
    public List<FuelTransactionRecord> DuplicatesSkipped { get; } = new();
    public List<string> Warnings { get; } = new();
    public decimal TotalImportedAmount => NewTransactions.Sum(t => t.Amount);
    public int TotalParsedCount => NewTransactions.Count + DuplicatesSkipped.Count;
}

public static class FuelImporter
{
    private static readonly Regex PlateRegex = new(@"\b([A-Z]{1,3})\s*(\d{1,4})\s*([A-Z]?)\b", RegexOptions.Compiled);

    public static FuelImportResult ProcessFile(
        string filePath,
        HashSet<string> existingTransactionSignatures,
        Dictionary<string, (string FullCard, string OwnerName, long OwnerId)> cardRegistry,
        Dictionary<string, string> vehicleTypeRegistry,
        Dictionary<string, long> staffRegistry)
    {
        var result = new FuelImportResult
        {
            SourceFileName = Path.GetFileName(filePath)
        };

        if (!File.Exists(filePath))
        {
            result.Warnings.Add($"File '{filePath}' does not exist.");
            return result;
        }

        string ext = Path.GetExtension(filePath).ToLowerInvariant();
        List<FuelTransactionRecord> rawRecords;

        if (ext == ".pdf")
        {
            rawRecords = ExtractFromPdf(filePath, result);
        }
        else if (ext == ".xlsx" || ext == ".xls" || ext == ".csv")
        {
            rawRecords = ExtractFromSpreadsheet(filePath, result);
        }
        else
        {
            result.Warnings.Add($"Unsupported file format '{ext}'. Please select a PDF, Excel, or CSV statement.");
            return result;
        }

        foreach (var rec in rawRecords)
        {
            // 1. Reconcile Card Number (especially for masked Shell cards e.g. 7077841******000208)
            rec.CardNumber = ResolveFullCardNumber(rec.CardNumber, rec.VehiclePlate, cardRegistry);

            // 2. Reconcile Vehicle Type from master registry if blank
            if (string.IsNullOrWhiteSpace(rec.VehicleType) && !string.IsNullOrWhiteSpace(rec.VehiclePlate))
            {
                string normPlate = NormalizePlate(rec.VehiclePlate);
                if (vehicleTypeRegistry.TryGetValue(normPlate, out string? vType))
                {
                    rec.VehicleType = vType;
                }
            }

            // 3. Reconcile Staff ID from Driver Name if missing
            if (rec.StaffID <= 0 && !string.IsNullOrWhiteSpace(rec.StaffName))
            {
                string normStaff = rec.StaffName.Trim().ToLowerInvariant();
                foreach (var (name, id) in staffRegistry)
                {
                    if (name.Trim().ToLowerInvariant().Contains(normStaff) || normStaff.Contains(name.Trim().ToLowerInvariant()))
                    {
                        rec.StaffID = id;
                        break;
                    }
                }
            }

            // 4. Duplicate Check Signature: Date + Time + NormalizedPlate + Amount
            string sig = $"{rec.Date:yyyyMMdd}_{rec.Time}_{NormalizePlate(rec.VehiclePlate)}_{rec.Amount:F2}";
            if (existingTransactionSignatures.Contains(sig))
            {
                result.DuplicatesSkipped.Add(rec);
            }
            else
            {
                result.NewTransactions.Add(rec);
                existingTransactionSignatures.Add(sig);
            }
        }

        return result;
    }

    private static List<FuelTransactionRecord> ExtractFromPdf(string pdfPath, FuelImportResult result)
    {
        var records = new List<FuelTransactionRecord>();
        using var document = PdfDocument.Open(pdfPath);

        var allPagesText = new List<string>();
        foreach (var page in document.GetPages())
        {
            allPagesText.Add(page.Text);
        }

        string fullContent = string.Join("\n", allPagesText);

        // Detect Provider
        if (fullContent.Contains("PETRONAS", StringComparison.OrdinalIgnoreCase) ||
            fullContent.Contains("smartPAY", StringComparison.OrdinalIgnoreCase))
        {
            result.Provider = FuelProvider.Petronas;
            records.AddRange(ParsePetronasPdf(document));
        }
        else if (fullContent.Contains("Shell Malaysia Trading", StringComparison.OrdinalIgnoreCase) ||
                 fullContent.Contains("fleethub.shell.com", StringComparison.OrdinalIgnoreCase))
        {
            result.Provider = FuelProvider.Shell;
            records.AddRange(ParseShellPdf(document));
        }
        else
        {
            result.Warnings.Add("Could not identify provider as either Petronas SmartPay or Shell Malaysia.");
        }

        return records;
    }

    /// <summary>
    /// Parses official Petronas SmartPay e-Statement (STM-*.pdf).
    /// </summary>
    private static List<FuelTransactionRecord> ParsePetronasPdf(PdfDocument document)
    {
        var list = new List<FuelTransactionRecord>();
        string currentVehiclePlate = string.Empty;
        string currentCardNumber = string.Empty;

        var headerCardRegex = new Regex(@"Nombor\s*Kad\s*[:\s]*(\d{16,19})", RegexOptions.IgnoreCase);
        var headerPlateRegex = new Regex(@"No\.\s*Pendaftaran\s*Kenderaan\s*[:\s]*([A-Z0-9\s]+)", RegexOptions.IgnoreCase);
        var txRegex = new Regex(
            @"(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2})\s+(\d{12})\s+(PSS\s+[^\d]+?)\s+([A-Z\s]+?)\s+(\d{6,10})\s+(Diesel[^\d]+|PRIMAX[^\d]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.,]+)\s+([\d\.,]+)",
            RegexOptions.IgnoreCase);

        foreach (var page in document.GetPages())
        {
            string pageText = page.Text;
            if (string.IsNullOrWhiteSpace(pageText) || !pageText.Contains("LAPORAN ANALISA KENDERAAN"))
                continue;

            var cardMatch = headerCardRegex.Match(pageText);
            if (cardMatch.Success) currentCardNumber = cardMatch.Groups[1].Value.Trim();

            var plateMatch = headerPlateRegex.Match(pageText);
            if (plateMatch.Success)
            {
                string rawP = plateMatch.Groups[1].Value.Trim();
                int newlineIdx = rawP.IndexOfAny(new[] { '\r', '\n' });
                if (newlineIdx > 0) rawP = rawP.Substring(0, newlineIdx).Trim();
                currentVehiclePlate = FormatPlate(rawP);
            }

            var lines = pageText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                var match = txRegex.Match(line);
                if (!match.Success) continue;

                if (!DateOnly.TryParseExact(match.Groups[1].Value, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var txDate))
                    continue;

                string timeFull = match.Groups[2].Value; // "HH:mm:ss"
                string timeClean = timeFull.Length >= 5 ? timeFull.Substring(0, 5) : timeFull;
                string station = match.Groups[4].Value.Trim();
                string driver = match.Groups[5].Value.Trim();
                long.TryParse(match.Groups[6].Value, out long odo);
                string product = match.Groups[7].Value.Trim();
                decimal.TryParse(match.Groups[8].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal qty);
                string netAmountStr = match.Groups[12].Value.Replace(",", "");
                decimal.TryParse(netAmountStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal netAmount);

                // Handle Spare cards where driver name field contains the actual vehicle plate
                string finalPlate = currentVehiclePlate;
                if ((finalPlate.Equals("SPARE", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(finalPlate)) &&
                    PlateRegex.IsMatch(driver))
                {
                    finalPlate = FormatPlate(driver);
                }

                list.Add(new FuelTransactionRecord
                {
                    Date = txDate,
                    Time = timeClean,
                    VehiclePlate = finalPlate,
                    CardNumber = currentCardNumber,
                    StaffName = driver,
                    StationName = station,
                    Odometer = odo,
                    Litres = qty,
                    ProductDescription = product,
                    Amount = netAmount,
                    ReceiptNumber = match.Groups[3].Value,
                    Provider = FuelProvider.Petronas
                });
            }
        }

        return list;
    }

    /// <summary>
    /// Parses official Shell Malaysia Trading e-Invoice (Invoice_*.pdf).
    /// </summary>
    private static List<FuelTransactionRecord> ParseShellPdf(PdfDocument document)
    {
        var list = new List<FuelTransactionRecord>();
        string currentCard = string.Empty;
        string currentDriver = string.Empty;
        string currentPlate = string.Empty;

        var cardHeaderRegex = new Regex(@"\b(7077\d{3}\*{6}\d{6})\s+([A-Z\s]+?)\s+([A-Z]{1,3}\s*\d{1,4}\s*[A-Z]?)\b");
        var txRowRegex = new Regex(
            @"(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})\s*(?:\d+)?\s*(SHELL\s+[A-Z0-9\s\.,/\-]+?)\s+(\d{5,8})\s+(\d+)\s+(\d{2})\s+([\d\.]+)\s*Lit\s+([\d\.]+)\s+([\d\.,]+)(?:\s+([\-\d\.,]+)\s+([\-\d\.,]+))?",
            RegexOptions.IgnoreCase);

        foreach (var page in document.GetPages())
        {
            string pageText = page.Text;
            if (!pageText.Contains("TRANSACTION DETAILS") && !pageText.Contains("BUTIRAN TRANSAKSI"))
                continue;

            var lines = pageText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                var cardMatch = cardHeaderRegex.Match(line);
                if (cardMatch.Success)
                {
                    currentCard = cardMatch.Groups[1].Value.Trim();
                    currentDriver = cardMatch.Groups[2].Value.Trim();
                    currentPlate = FormatPlate(cardMatch.Groups[3].Value.Trim());
                    continue;
                }

                var txMatch = txRowRegex.Match(line);
                if (!txMatch.Success) continue;

                if (!DateOnly.TryParseExact(txMatch.Groups[1].Value, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var txDate))
                    continue;

                string time = txMatch.Groups[2].Value;
                string station = CleanStationName(txMatch.Groups[3].Value);
                string receipt = txMatch.Groups[4].Value;
                long.TryParse(txMatch.Groups[5].Value, out long odo);
                decimal.TryParse(txMatch.Groups[7].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal qty);
                decimal.TryParse(txMatch.Groups[9].Value.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal grossAmount);

                decimal subsidyAmount = 0m;
                if (txMatch.Groups[10].Success && !string.IsNullOrWhiteSpace(txMatch.Groups[10].Value))
                {
                    decimal.TryParse(txMatch.Groups[10].Value.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out subsidyAmount);
                }

                // Net company cost after government subsidy credit
                decimal netAmount = grossAmount + (subsidyAmount < 0 ? subsidyAmount : 0m);

                list.Add(new FuelTransactionRecord
                {
                    Date = txDate,
                    Time = time,
                    VehiclePlate = currentPlate,
                    CardNumber = currentCard,
                    StaffName = currentDriver,
                    StationName = station,
                    Odometer = odo,
                    Litres = qty,
                    Amount = netAmount,
                    ReceiptNumber = receipt,
                    ProductDescription = $"Shell Code {txMatch.Groups[6].Value}",
                    Provider = FuelProvider.Shell
                });
            }
        }

        return list;
    }

    /// <summary>
    /// Fallback for portal spreadsheet downloads (.xlsx, .xls, .csv).
    /// </summary>
    private static List<FuelTransactionRecord> ExtractFromSpreadsheet(string path, FuelImportResult result)
    {
        var records = new List<FuelTransactionRecord>();
        var lines = File.ReadAllLines(path);
        if (lines.Length == 0) return records;

        // Sniff CSV or delimited text
        char delim = lines[0].Contains('\t') ? '\t' : ',';
        int headerRow = -1;
        var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

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

        if (headerRow < 0) return records;

        var headerCols = lines[headerRow].Split(delim);
        for (int c = 0; c < headerCols.Length; c++)
        {
            string clean = headerCols[c].Trim().Replace("\"", "").ToLowerInvariant();
            colMap[clean] = c;
        }

        for (int i = headerRow + 1; i < lines.Length; i++)
        {
            string rawLine = lines[i].Trim();
            if (string.IsNullOrWhiteSpace(rawLine) || rawLine.Contains("TOTAL")) continue;

            var parts = rawLine.Split(delim);
            string dateStr = GetCol(parts, colMap, "date");
            if (string.IsNullOrWhiteSpace(dateStr)) continue;

            if (DateTime.TryParse(dateStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt) ||
                DateTime.TryParse(dateStr, new CultureInfo("en-GB"), DateTimeStyles.None, out dt))
            {
                decimal.TryParse(GetCol(parts, colMap, "amount").Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amt);
                long.TryParse(GetCol(parts, colMap, "odometer"), out long odo);

                records.Add(new FuelTransactionRecord
                {
                    Date = DateOnly.FromDateTime(dt),
                    Time = dt.ToString("HH:mm"),
                    CardNumber = GetCol(parts, colMap, "card").Replace("'", ""),
                    VehiclePlate = FormatPlate(GetCol(parts, colMap, "vehicle")),
                    StaffName = GetCol(parts, colMap, "driver"),
                    StationName = GetCol(parts, colMap, "station"),
                    Amount = amt,
                    Odometer = odo
                });
            }
        }

        return records;
    }

    private static string ResolveFullCardNumber(
        string rawCard,
        string vehiclePlate,
        Dictionary<string, (string FullCard, string OwnerName, long OwnerId)> cardRegistry)
    {
        if (string.IsNullOrWhiteSpace(rawCard)) return string.Empty;
        string clean = rawCard.Replace("'", "").Trim();

        // Direct exact match
        if (cardRegistry.TryGetValue(clean, out var exact))
            return exact.FullCard;

        // Shell Masked Card Pattern (e.g. 7077841******000208)
        if (clean.Contains("*"))
        {
            string prefix = clean.Substring(0, clean.IndexOf('*'));
            string suffix = clean.Substring(clean.LastIndexOf('*') + 1);

            foreach (var (fullNum, val) in cardRegistry)
            {
                if (fullNum.StartsWith(prefix) && fullNum.EndsWith(suffix))
                {
                    return fullNum;
                }
            }
        }

        return clean;
    }

    public static string FormatPlate(string rawPlate)
    {
        if (string.IsNullOrWhiteSpace(rawPlate)) return string.Empty;
        string clean = Regex.Replace(rawPlate.Trim(), @"\s+", "");
        var m = PlateRegex.Match(clean);
        if (m.Success)
        {
            return $"{m.Groups[1].Value} {m.Groups[2].Value} {m.Groups[3].Value}".Trim();
        }
        return clean.ToUpperInvariant();
    }

    public static string NormalizePlate(string plate) =>
        Regex.Replace(plate ?? string.Empty, @"[\s_\-]+", "").ToUpperInvariant();

    private static string CleanStationName(string raw)
    {
        string s = Regex.Replace(raw, @"\s+", " ").Trim();
        int slashIdx = s.IndexOf('/');
        if (slashIdx > 0) s = s.Substring(0, slashIdx).Trim();
        return s;
    }

    private static string GetCol(string[] parts, Dictionary<string, int> map, string keyword)
    {
        foreach (var (k, idx) in map)
        {
            if (k.Contains(keyword) && idx < parts.Length)
            {
                return parts[idx].Trim().Replace("\"", "");
            }
        }
        return string.Empty;
    }
}