namespace SKSB_App.Core.ERPBridge;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public static class ERPExporter
{
    /// <summary>
    /// Writes raw header and row values directly to a delimited file on disk.
    /// </summary>
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

    /// <summary>
    /// Exports strongly typed collections to a delimited file via a custom row mapping function.
    /// </summary>
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

        // Write Headers
        sb.AppendLine(string.Join(delimiter, headers.Select(h => ERPHelper.EscapeDelimitedCell(h, delimiter))));

        // Write Rows
        foreach (var row in rows)
        {
            sb.AppendLine(string.Join(delimiter, row.Select(cell => ERPHelper.EscapeDelimitedCell(cell, delimiter))));
        }

        return sb.ToString();
    }
}