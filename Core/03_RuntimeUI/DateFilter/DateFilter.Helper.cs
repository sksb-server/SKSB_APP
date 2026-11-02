namespace SKSB_App.Core.RuntimeUI.DateFilter;

using System;
using System.Collections.Generic;
using System.Linq;

public static class DateFilterHelper
{
    public static bool IsWithinScope(this DateFilterScope scope, DateOnly targetDate, DateOnly anchorDate)
    {
        return scope switch
        {
            DateFilterScope.All_Time or DateFilterScope.None => true,
            DateFilterScope.Annual => targetDate.Year == anchorDate.Year,
            DateFilterScope.Month => targetDate.Year == anchorDate.Year && targetDate.Month == anchorDate.Month,
            DateFilterScope.Day => targetDate == anchorDate,
            _ => true
        };
    }

    public static DateFilterScope ParseScope(string? scopeStr)
    {
        if (string.IsNullOrWhiteSpace(scopeStr))
            return DateFilterScope.All_Time;

        string clean = scopeStr.Trim().Replace(" ", "_");
        if (Enum.TryParse<DateFilterScope>(clean, true, out var scope))
            return scope == DateFilterScope.None ? DateFilterScope.All_Time : scope;

        return clean.ToLowerInvariant() switch
        {
            "year" or "yearly" => DateFilterScope.Annual,
            "monthly" => DateFilterScope.Month,
            "daily" or "today" => DateFilterScope.Day,
            _ => DateFilterScope.All_Time
        };
    }

    public static IEnumerable<T> FilterByDateScope<T>(
        this IEnumerable<T>? items,
        Func<T, DateOnly> dateSelector,
        DateOnly anchorDate,
        DateFilterScope scope)
    {
        if (items == null) return Enumerable.Empty<T>();
        return items.Where(item => scope.IsWithinScope(dateSelector(item), anchorDate));
    }

    public static decimal SumByDateScope<T>(
        this IEnumerable<T>? items,
        Func<T, DateOnly> dateSelector,
        Func<T, decimal> amountSelector,
        DateOnly anchorDate,
        DateFilterScope scope)
    {
        if (items == null) return 0.00m;
        return items
            .Where(item => scope.IsWithinScope(dateSelector(item), anchorDate))
            .Sum(amountSelector);
    }
}