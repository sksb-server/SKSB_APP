namespace SKSB_App.Core.Integration.Erp;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public static class ErpExporter
{
    public static void ExportToFile(
        string filePath,
        IEnumerable<string> headers,
        IEnumerable<IEnumerable<object?>> rows,
        char delimiter = ',',
        Encoding? encoding = null)
    {
        string content = BuildDelimitedContent(headers, rows, delimiter);
        string? directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
        File.WriteAllText(filePath, content, encoding ?? Encoding.UTF8);
    }

    public static void ExportEntitiesToFile<T>(
        string filePath,
        IEnumerable<string> headers,
        IEnumerable<T> records,
        Func<T, IEnumerable<object?>> rowMapper,
        char delimiter = ',',
        Encoding? encoding = null)
    {
        var rows = new List<IEnumerable<object?>>();
        foreach (var record in records)
        {
            rows.Add(rowMapper(record));
        }
        ExportToFile(filePath, headers, rows, delimiter, encoding);
    }

    public static string BuildDelimitedContent(
        IEnumerable<string> headers,
        IEnumerable<IEnumerable<object?>> rows,
        char delimiter = ',')
    {
        var sb = new StringBuilder();

        sb.AppendLine(string.Join(delimiter, headers.Select(h => ErpHelper.EscapeDelimitedCell(h, delimiter))));

        foreach (var row in rows)
        {
            sb.AppendLine(string.Join(delimiter, row.Select(cell => ErpHelper.EscapeDelimitedCell(cell, delimiter))));
        }

        return sb.ToString();
    }
}