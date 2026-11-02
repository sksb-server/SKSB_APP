namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using SKSB_App.Core.Integration.Erp;

public class FuelImportResult
{
    public List<FuelImportRow> ImportedTransactions { get; set; } = new();
    public List<ImportSkippedRecord> SkippedRecords { get; set; } = new();
    public int TotalRowsProcessed { get; set; } = 0;
    public string ProviderDetected { get; set; } = "Unknown";

    public ImportSummaryReport ToGenericReport() => new()
    {
        SourceSystemName = $"{ProviderDetected} Fleet Card",
        EntityLabel = "fuel transaction(s)",
        SuccessfullyImportedCount = ImportedTransactions.Count,
        FilteredExclusionCount = 0,
        FilterRuleDescription = "Mapped against active FuelCardsTable. Auto-derived Odometer Before and Card Owner.",
        SkippedErrors = SkippedRecords
    };
}

public class FuelImportRow
{
    public DateOnly Date { get; set; }
    public string Time { get; set; } = string.Empty;
    public string CardNumber { get; set; } = string.Empty;
    public string VehiclePlate { get; set; } = string.Empty;
    public string VehicleType { get; set; } = string.Empty;
    public long StaffID { get; set; }
    public string StaffName { get; set; } = string.Empty;
    public string StationName { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public int Odometer { get; set; }
}

public static class FuelImporter
{
    public static FuelImportResult ImportTransactions(
        string filePath,
        Dictionary<string, (string Plate, string Type, long StaffId, string StaffName)> cardRegistry)
    {
        var result = new FuelImportResult();
        var doc = ErpImporter.LoadFromFile(filePath);
        if (doc.Rows.Count == 0)
            return result;

        // Auto-detect Provider based on column headers and card patterns
        bool isShell = doc.HasColumn("Vehicle Registration") || doc.HasColumn("VRN") || doc.HasColumn("Shell Card Number");
        result.ProviderDetected = isShell ? "Shell Fleet Hub" : "Petronas SmartPay";

        for (int r = 0; r < doc.Rows.Count; r++)
        {
            var row = doc.Rows[r];
            int excelRowNum = r + 2;
            result.TotalRowsProcessed++;

            // 1. Extract & Sanitize Card Number
            string rawCard = isShell
                ? (row.GetString("Card Number") ?? row.GetString("Shell Card Number"))
                : row.GetString("Card Number");
            string cleanCard = rawCard.Replace("'", "").Replace(" ", "").Trim();

            if (string.IsNullOrWhiteSpace(cleanCard) || cleanCard.Contains("TRANSACTION") || cleanCard.Contains("UNBILL"))
            {
                continue;
            }

            // 2. Extract Date & Time
            string rawDateTimeStr = isShell
                ? (row.GetString("Transaction Date & Time") ?? row.GetString("Transaction Date") ?? row.GetString("Date"))
                : (row.GetString("Date Time") ?? row.GetString("Date / Time") ?? row.GetString("Date"));

            DateTime parsedDt;
            if (!DateTime.TryParse(rawDateTimeStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDt))
            {
                if (!DateTime.TryParse(rawDateTimeStr, out parsedDt))
                {
                    result.SkippedRecords.Add(new ImportSkippedRecord
                    {
                        ExcelRowNumber = excelRowNum,
                        Identifier = cleanCard,
                        NameOrTitle = "Invalid Date",
                        Reason = $"Could not parse transaction date/time '{rawDateTimeStr}'."
                    });
                    continue;
                }
            }

            DateOnly txDate = DateOnly.FromDateTime(parsedDt);
            string txTime = parsedDt.ToString("HH:mm");

            // 3. Extract Amount
            string rawAmount = isShell
                ? (row.GetString("Gross Amount (MYR)") ?? row.GetString("Gross Amount") ?? row.GetString("Amount"))
                : (row.GetString("Transaction Amount (RM)") ?? row.GetString("Amount"));

            decimal.TryParse(rawAmount?.Replace("RM", "").Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amount);

            // 4. Resolve Vehicle & Assigned Staff via FuelCardsTable registry
            string plate = string.Empty;
            string vType = string.Empty;
            long staffId = 0;
            string staffName = string.Empty;

            if (cardRegistry.TryGetValue(cleanCard, out var cardInfo))
            {
                plate = cardInfo.Plate;
                vType = cardInfo.Type;
                staffId = cardInfo.StaffId;
                staffName = cardInfo.StaffName;
            }
            else
            {
                plate = row.GetString("Vehicle Registration") ?? row.GetString("Vehicle No.") ?? row.GetString("VRN");
                staffName = row.GetString("Driver") ?? row.GetString("Driver Name");
            }

            // 5. Extract Station & Odometer
            string station = isShell
                ? (row.GetString("Station Name") ?? row.GetString("Site Name") ?? row.GetString("Location"))
                : (row.GetString("Station Name") ?? row.GetString("Site"));

            string rawOdo = row.GetString("Odometer") ?? row.GetString("Mileage");
            int.TryParse(rawOdo?.Replace(",", "").Trim(), out int odometer);

            result.ImportedTransactions.Add(new FuelImportRow
            {
                Date = txDate,
                Time = txTime,
                CardNumber = cleanCard,
                VehiclePlate = plate,
                VehicleType = vType,
                StaffID = staffId,
                StaffName = staffName,
                StationName = station,
                Amount = amount,
                Odometer = odometer
            });
        }

        return result;
    }
}