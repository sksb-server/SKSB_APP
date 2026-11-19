namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Collections.Generic;

/// <summary>
/// Natural (numeric-aware), case-insensitive string comparer. Digit runs compare by numeric value,
/// other characters compare case-insensitively; ordinal comparison is the final tie-breaker.
/// </summary>
public sealed class NaturalStringComparer : IComparer<string?>
{
    public static readonly NaturalStringComparer Instance = new();

    private NaturalStringComparer() { }

    public int Compare(string? x, string? y)
    {
        if (ReferenceEquals(x, y)) return 0;
        if (x is null) return -1;
        if (y is null) return 1;

        ReadOnlySpan<char> sx = x.AsSpan().Trim();
        ReadOnlySpan<char> sy = y.AsSpan().Trim();
        int ix = 0, iy = 0;

        while (ix < sx.Length && iy < sy.Length)
        {
            if (char.IsDigit(sx[ix]) && char.IsDigit(sy[iy]))
            {
                int nx = ix; while (nx < sx.Length && char.IsDigit(sx[nx])) nx++;
                int ny = iy; while (ny < sy.Length && char.IsDigit(sy[ny])) ny++;

                ReadOnlySpan<char> dx = sx.Slice(ix, nx - ix).TrimStart('0');
                ReadOnlySpan<char> dy = sy.Slice(iy, ny - iy).TrimStart('0');

                int c = dx.Length.CompareTo(dy.Length);
                if (c != 0) return c;
                c = dx.SequenceCompareTo(dy);
                if (c != 0) return c;

                ix = nx; iy = ny;
            }
            else
            {
                int c = char.ToUpperInvariant(sx[ix]).CompareTo(char.ToUpperInvariant(sy[iy]));
                if (c != 0) return c;
                ix++; iy++;
            }
        }

        if (ix < sx.Length) return 1;
        if (iy < sy.Length) return -1;

        return StringComparer.Ordinal.Compare(x, y);
    }
}
