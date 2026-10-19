namespace SKSB_App.Core.ErpBridge;

using ExcelDna.Integration;
using SKSB_App.Core.ERPBridge;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class ERPTableDocument
{
    public List<string> Headers { get; }
    public List<ERPTableRow> Rows { get; }

    private readonly Dictionary<string, int> _headerMap = new(StringComparer.OrdinalIgnoreCase);

    public ERPTableDocument(List<string> headers, List<ERPTableRow> rows)
    {
        Headers = headers;
        Rows = rows;

        for (int i = 0; i < headers.Count; i++)
        {
            string clean = ERPHelper.NormalizeHeader(headers[i]);
            if (!_headerMap.ContainsKey(clean))
            {
                _headerMap[clean] = i;
            }
        }
    }

    public int GetColumnIndex(string columnName)
    {
        string target = ERPHelper.NormalizeHeader(columnName);
        return _headerMap.TryGetValue(target, out int index) ? index : -1;
    }

    public bool HasColumn(string columnName) => GetColumnIndex(columnName) >= 0;
}

public class ERPTableRow
{
    private readonly string[] _cells;
    private readonly ERPTableDocument _owner;

    public ERPTableRow(string[] cells, ERPTableDocument owner)
    {
        _cells = cells;
        _owner = owner;
    }

    public string GetString(string columnName, string defaultValue = "")
    {
        int idx = _owner.GetColumnIndex(columnName);
        if (idx < 0 || idx >= _cells.Length) return defaultValue;
        return _cells[idx].Trim();
    }

    public long GetLong(string columnName, long defaultValue = 0) =>
        ERPHelper.ParseLong(GetString(columnName), defaultValue);

    public decimal GetDecimal(string columnName, decimal defaultValue = 0.00m) =>
        ERPHelper.ParseDecimal(GetString(columnName), defaultValue);

    public DateOnly? GetDateOnly(string columnName) =>
        ERPHelper.ParseDateOnly(GetString(columnName));

    public bool GetBool(string columnName, bool defaultValue = false)
    {
        string val = GetString(columnName);
        if (string.IsNullOrWhiteSpace(val)) return defaultValue;
        if (bool.TryParse(val, out bool b)) return b;
        if (val.Equals("1") || val.Equals("YES", StringComparison.OrdinalIgnoreCase)) return true;
        if (val.Equals("0") || val.Equals("NO", StringComparison.OrdinalIgnoreCase)) return false;
        return defaultValue;
    }
}

public static class ERPImporter
{
    public static ERPTableDocument LoadFromFile(string filePath, char delimiter = ',')
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"ERP import file not found at: {filePath}");

        string ext = Path.GetExtension(filePath).ToLowerInvariant();
        if (ext == ".xlsx" || ext == ".xls")
        {
            return LoadFromExcelSpreadsheet(filePath);
        }

        var lines = File.ReadAllLines(filePath);
        return LoadFromLines(lines, delimiter);
    }

    private static ERPTableDocument LoadFromExcelSpreadsheet(string filePath)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = null;
        try
        {
            // Open read-only in background
            wb = xlApp.Workbooks.Open(filePath, ReadOnly: true, AddToMru: false);
            dynamic ws = wb.Sheets[1];
            dynamic usedRange = ws.UsedRange;
            object[,] data = usedRange.Value2;

            if (data == null)
                return new ERPTableDocument(new List<string>(), new List<ERPTableRow>());

            int rowCount = data.GetLength(0);
            int colCount = data.GetLength(1);

            if (rowCount < 1)
                return new ERPTableDocument(new List<string>(), new List<ERPTableRow>());

            // Row 1: Headers
            var headers = new List<string>();
            for (int c = 1; c <= colCount; c++)
            {
                headers.Add(data[1, c]?.ToString()?.Trim() ?? string.Empty);
            }

            var rows = new List<ERPTableRow>();
            var doc = new ERPTableDocument(headers, rows);

            // Rows 2..N: Data
            for (int r = 2; r <= rowCount; r++)
            {
                var cells = new string[colCount];
                bool hasValue = false;
                for (int c = 1; c <= colCount; c++)
                {
                    string cellVal = data[r, c]?.ToString()?.Trim() ?? string.Empty;
                    cells[c - 1] = cellVal;
                    if (!string.IsNullOrWhiteSpace(cellVal)) hasValue = true;
                }

                if (hasValue)
                {
                    rows.Add(new ERPTableRow(cells, doc));
                }
            }

            return doc;
        }
        finally
        {
            wb?.Close(SaveChanges: false);
        }
    }

    private static ERPTableDocument LoadFromLines(string[] lines, char delimiter)
    {
        var validLines = lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
        if (validLines.Count == 0)
            return new ERPTableDocument(new List<string>(), new List<ERPTableRow>());

        var headers = ERPHelper.SplitDelimitedLine(validLines[0], delimiter).ToList();
        var rows = new List<ERPTableRow>();
        var document = new ERPTableDocument(headers, rows);

        for (int i = 1; i < validLines.Count; i++)
        {
            var cells = ERPHelper.SplitDelimitedLine(validLines[i], delimiter);
            rows.Add(new ERPTableRow(cells, document));
        }

        return document;
    }
}