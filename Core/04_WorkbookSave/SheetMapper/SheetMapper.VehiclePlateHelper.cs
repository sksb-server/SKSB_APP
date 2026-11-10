namespace SKSB_App.Core.WorkbookSave.SheetMapper;

using System.Text.RegularExpressions;

/// <summary>
/// Single implementation of Malaysian vehicle plate detection, normalisation and display formatting,
/// shared by every module that reads or writes plates.
/// </summary>
public static class VehiclePlateHelper
{
    /// <summary>Finds a plate inside free text, e.g. "WXY 1234 A" within a statement line.</summary>
    public static readonly Regex PlateRegex = new(@"\b([A-Z]{1,3})\s*(\d{1,4})\s*([A-Z]?)\b", RegexOptions.Compiled);

    private static readonly Regex SeparatorRegex = new(@"[\s_\-]+", RegexOptions.Compiled);
    private static readonly Regex CompactPlateRegex = new(@"^([A-Z]{1,3})(\d{1,4})([A-Z]?)$", RegexOptions.Compiled);

    /// <summary>Comparison form: separators stripped, upper-cased ("wxy-1234 a" -> "WXY1234A").</summary>
    public static string NormalizePlate(string? plate) =>
        SeparatorRegex.Replace(plate ?? string.Empty, "").ToUpperInvariant();

    /// <summary>Display form: "WXY 1234 A". Unrecognised patterns are returned normalised.</summary>
    public static string FormatPlate(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        string clean = NormalizePlate(raw);
        var m = CompactPlateRegex.Match(clean);
        if (m.Success)
        {
            return $"{m.Groups[1].Value} {m.Groups[2].Value} {m.Groups[3].Value}".Trim();
        }
        return clean;
    }
}
