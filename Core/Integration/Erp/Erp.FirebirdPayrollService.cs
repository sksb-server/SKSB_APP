namespace SKSB_App.Core.Integration.Erp;

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using FirebirdSql.Data.FirebirdClient;
using SKSB_App.Core.Diagnostics;

// =========================================================================
// 1. PAYROLL DTOs (read-only projection of SQL Payroll HR tables)
// =========================================================================

/// <summary>One row of SQL Payroll <c>HR_EMP</c>.</summary>
public sealed record PayrollEmployeeRow(
    string Code,
    string Name,
    string Department,
    string HrGroup,
    string JobTitle,
    string NewIc,
    string Passport,
    string Mobile,
    string Phone1,
    DateOnly? JoinDate,
    DateOnly? ResignDate,
    bool IsActive,
    string WagesType,
    decimal Wages);

/// <summary>One "Revise Wages..." row of SQL Payroll <c>HR_EMP_HISTORY</c>.</summary>
public sealed record PayrollWageHistoryRow(long AutoKey, string Code, DateOnly PostDate, string Description);

/// <summary>Everything read from SQL Payroll in a single read-only session.</summary>
public sealed record PayrollSnapshot(List<PayrollEmployeeRow> Employees, List<PayrollWageHistoryRow> WageHistory, string ConnectedHost)
{
    /// <summary>Every non-voided LV_APP leave row (read in the same Firebird session as employees + wage history).</summary>
    public List<PayrollLeaveRow> Leaves { get; init; } = new();
}

/// <summary>One public holiday of SQL Payroll <c>LV_CALENDAR_DAYS</c> (<c>DAYTYPE = 2</c>).</summary>
public sealed record PayrollHolidayRow(long AutoKey, DateOnly HolidayDate, string Description);

/// <summary>One leave application day of SQL Payroll <c>LV_APP</c> (<c>CODE</c> = employee code = Staff ID).</summary>
public sealed record PayrollLeaveRow(long DocKey, string Code, DateOnly LeaveDate, string LeaveType, decimal LeaveDay, string Description);

/// <summary>
/// Read-only access to the SKSB SQL Payroll Firebird database (<c>PAY-0003.FDB</c>). Every query runs inside a
/// read-only, read-committed transaction; this service never writes to SQL Payroll.
/// </summary>
public class FirebirdPayrollService
{
    // =========================================================================
    // 2. EMBEDDED READ-ONLY CONNECTION CONFIG
    // =========================================================================
    public const string PrimaryHost = "LAPTOP-MPMUNNTG";
    public const string FallbackHost = "192.168.0.219";
    public const int Port = 3050;
    public const string DatabasePath = @"C:\eStream\SQL Payroll\DB\PAY-0003.FDB";
    public const string UserId = "SKSB_RO";
    public const string Password = "2QCgCfj1lyUTvKKgJ4kA";
    public const int ConnectionTimeoutSeconds = 5;
    public const int CommandTimeoutSeconds = 10;

    private const string SqlEmployees =
        "SELECT CODE, NAME, DEPARTMENT, HRGROUP, JOBTITLE, NEWIC, PASSPORT, MOBILE, PHONE1, JOINDATE, RESIGNDATE, ISACTIVE, WAGESTYPE, WAGES FROM HR_EMP ORDER BY CODE";

    private const string SqlWageHistory =
        "SELECT AUTOKEY, CODE, POSTDATE, DESCRIPTION FROM HR_EMP_HISTORY WHERE DESCRIPTION LIKE 'Revise Wages%' ORDER BY CODE, POSTDATE, AUTOKEY";

    private const string SqlPublicHolidays =
        "SELECT AUTOKEY, CALDAY, DESCRIPTION FROM LV_CALENDAR_DAYS WHERE DAYTYPE = 2 ORDER BY CALDAY DESC";

    /// <summary>LEAVETYPE '删' ("deleted") marks voided applications in SQL Payroll.</summary>
    private const string SqlLeavesForDate =
        "SELECT DOCKEY, CODE, LEAVEDATE, LEAVETYPE, LEAVEDAY, DESCRIPTION FROM LV_APP " +
        "WHERE LEAVEDATE = @LeaveDate AND COALESCE(LEAVETYPE, '') <> '删' ORDER BY CODE, DOCKEY";

    private const string SqlAllLeaves =
        "SELECT DOCKEY, CODE, LEAVEDATE, LEAVETYPE, LEAVEDAY, DESCRIPTION FROM LV_APP " +
        "WHERE COALESCE(LEAVETYPE, '') <> '删' ORDER BY LEAVEDATE DESC, DOCKEY DESC";

    /// <summary>Full LV_APP scans can be large (20k+ rows); give them more headroom than point queries.</summary>
    private const int LeaveScanTimeoutSeconds = 60;

    public static string BuildConnectionString(string host) => new FbConnectionStringBuilder
    {
        DataSource = host,
        Port = Port,
        Database = DatabasePath,
        UserID = UserId,
        Password = Password,
        Charset = "UTF8",
        ConnectionTimeout = ConnectionTimeoutSeconds,
        Pooling = false,
        Dialect = 3
    }.ToString();

    // =========================================================================
    // 3. SNAPSHOT FETCH (primary host -> fallback host)
    // =========================================================================
    /// <summary>Reads HR_EMP + wage revision history, trying the primary host first and then the fallback IP.</summary>
    public PayrollSnapshot FetchPayrollSnapshot()
    {
        var failures = new List<string>();
        foreach (string host in new[] { PrimaryHost, FallbackHost })
        {
            try
            {
                var snapshot = FetchFromHost(host);
                AppLogger.Info($"[SQL Payroll] Read {snapshot.Employees.Count} employee(s), {snapshot.WageHistory.Count} wage revision(s), {snapshot.Leaves.Count} leave row(s) from {host}.");
                return snapshot;
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[SQL Payroll] Host '{host}' failed: {ex.Message}");
                failures.Add($"{host}: {ex.Message}");
            }
        }
        throw new InvalidOperationException("SQL Payroll database is unreachable (" + string.Join(" | ", failures) + ").");
    }

    private static PayrollSnapshot FetchFromHost(string host)
    {
        using var conn = new FbConnection(BuildConnectionString(host));
        conn.Open();

        using var trans = conn.BeginTransaction(new FbTransactionOptions
        {
            TransactionBehavior = FbTransactionBehavior.Read | FbTransactionBehavior.ReadCommitted | FbTransactionBehavior.RecVersion
        });

        var employees = new List<PayrollEmployeeRow>();
        using (var cmd = new FbCommand(SqlEmployees, conn, trans) { CommandTimeout = CommandTimeoutSeconds })
        using (var reader = cmd.ExecuteReader())
        {
            while (reader.Read())
            {
                employees.Add(new PayrollEmployeeRow(
                    Code: Str(reader, 0),
                    Name: Str(reader, 1),
                    Department: Str(reader, 2),
                    HrGroup: Str(reader, 3),
                    JobTitle: Str(reader, 4),
                    NewIc: Str(reader, 5),
                    Passport: Str(reader, 6),
                    Mobile: Str(reader, 7),
                    Phone1: Str(reader, 8),
                    JoinDate: Date(reader, 9),
                    ResignDate: Date(reader, 10),
                    IsActive: Bool(reader, 11),
                    WagesType: Str(reader, 12),
                    Wages: Dec(reader, 13)));
            }
        }

        var history = new List<PayrollWageHistoryRow>();
        using (var cmd = new FbCommand(SqlWageHistory, conn, trans) { CommandTimeout = CommandTimeoutSeconds })
        using (var reader = cmd.ExecuteReader())
        {
            while (reader.Read())
            {
                DateOnly? post = Date(reader, 2);
                if (!post.HasValue) continue;
                long autoKey = reader.IsDBNull(0) ? 0 : Convert.ToInt64(reader.GetValue(0), CultureInfo.InvariantCulture);
                history.Add(new PayrollWageHistoryRow(autoKey, Str(reader, 1), post.Value, Str(reader, 3)));
            }
        }

        var leaves = ReadAllLeaves(conn, trans);

        // Read-only transaction: nothing to commit, roll back to release the snapshot cleanly.
        try { trans.Rollback(); } catch { }
        return new PayrollSnapshot(employees, history, host) { Leaves = leaves };
    }

    private static List<PayrollLeaveRow> ReadAllLeaves(FbConnection conn, FbTransaction trans)
    {
        var list = new List<PayrollLeaveRow>();
        using var cmd = new FbCommand(SqlAllLeaves, conn, trans) { CommandTimeout = LeaveScanTimeoutSeconds };
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            DateOnly? day = Date(reader, 2);
            if (!day.HasValue) continue;
            list.Add(new PayrollLeaveRow(
                DocKey: Long(reader, 0),
                Code: Str(reader, 1),
                LeaveDate: day.Value,
                LeaveType: Str(reader, 3),
                LeaveDay: reader.IsDBNull(4) ? 1m : Dec(reader, 4),
                Description: Str(reader, 5)));
        }
        return list;
    }

    // =========================================================================
    // 4. LEAVE CALENDAR (public holidays + date-scoped leave applications)
    // =========================================================================
    /// <summary>Every public holiday (<c>LV_CALENDAR_DAYS.DAYTYPE = 2</c>), newest first. Throws when unreachable.</summary>
    public List<PayrollHolidayRow> FetchPublicHolidays()
    {
        return RunReadOnly("public holidays", (conn, trans) =>
        {
            var list = new List<PayrollHolidayRow>();
            using var cmd = new FbCommand(SqlPublicHolidays, conn, trans) { CommandTimeout = CommandTimeoutSeconds };
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                DateOnly? day = Date(reader, 1);
                if (!day.HasValue) continue;
                list.Add(new PayrollHolidayRow(Long(reader, 0), day.Value, Str(reader, 2)));
            }
            return list;
        });
    }

    /// <summary>Non-voided leave applications (<c>LV_APP</c>) on <paramref name="date"/> only. Throws when unreachable.</summary>
    public List<PayrollLeaveRow> FetchLeavesForDate(DateOnly date)
    {
        return RunReadOnly($"leaves {date:yyyy-MM-dd}", (conn, trans) =>
        {
            var list = new List<PayrollLeaveRow>();
            using var cmd = new FbCommand(SqlLeavesForDate, conn, trans) { CommandTimeout = CommandTimeoutSeconds };
            cmd.Parameters.Add(new FbParameter("@LeaveDate", FbDbType.Date) { Value = date.ToDateTime(TimeOnly.MinValue) });
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(new PayrollLeaveRow(
                    DocKey: Long(reader, 0),
                    Code: Str(reader, 1),
                    LeaveDate: Date(reader, 2) ?? date,
                    LeaveType: Str(reader, 3),
                    LeaveDay: reader.IsDBNull(4) ? 1m : Dec(reader, 4),
                    Description: Str(reader, 5)));
            }
            return list;
        });
    }

    /// <summary>Every non-voided leave application (<c>LV_APP</c>), newest first. Throws when unreachable.</summary>
    public List<PayrollLeaveRow> FetchAllLeaves() => RunReadOnly("all leaves", ReadAllLeaves);

    /// <summary>Runs <paramref name="body"/> in a read-only transaction on the primary host, then the fallback IP.</summary>
    private static T RunReadOnly<T>(string label, Func<FbConnection, FbTransaction, T> body)
    {
        var failures = new List<string>();
        foreach (string host in new[] { PrimaryHost, FallbackHost })
        {
            try
            {
                using var conn = new FbConnection(BuildConnectionString(host));
                conn.Open();
                using var trans = conn.BeginTransaction(new FbTransactionOptions
                {
                    TransactionBehavior = FbTransactionBehavior.Read | FbTransactionBehavior.ReadCommitted | FbTransactionBehavior.RecVersion
                });

                T result = body(conn, trans);
                try { trans.Rollback(); } catch { }
                AppLogger.Info($"[SQL Payroll] Read {label} from {host}.");
                return result;
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[SQL Payroll] Host '{host}' failed reading {label}: {ex.Message}");
                failures.Add($"{host}: {ex.Message}");
            }
        }
        throw new InvalidOperationException("SQL Payroll database is unreachable (" + string.Join(" | ", failures) + ").");
    }

    // =========================================================================
    // 5. READER HELPERS
    // =========================================================================
    private static long Long(IDataRecord r, int i)
    {
        if (r.IsDBNull(i)) return 0;
        try { return Convert.ToInt64(r.GetValue(i), CultureInfo.InvariantCulture); }
        catch { return 0; }
    }

    private static string Str(IDataRecord r, int i) =>
        r.IsDBNull(i) ? string.Empty : (Convert.ToString(r.GetValue(i), CultureInfo.InvariantCulture) ?? string.Empty).Trim();

    private static DateOnly? Date(IDataRecord r, int i)
    {
        if (r.IsDBNull(i)) return null;
        object v = r.GetValue(i);
        return v switch
        {
            DateTime dt when dt.Year > 1900 => DateOnly.FromDateTime(dt),
            DateOnly d when d.Year > 1900 => d,
            _ => null
        };
    }

    private static decimal Dec(IDataRecord r, int i)
    {
        if (r.IsDBNull(i)) return 0m;
        try { return Convert.ToDecimal(r.GetValue(i), CultureInfo.InvariantCulture); }
        catch { return 0m; }
    }

    /// <summary>SQL Payroll booleans are stored as 'T'/'F' chars (or 1/0 smallints, depending on version).</summary>
    private static bool Bool(IDataRecord r, int i)
    {
        if (r.IsDBNull(i)) return false;
        object v = r.GetValue(i);
        if (v is bool b) return b;
        string s = (Convert.ToString(v, CultureInfo.InvariantCulture) ?? string.Empty).Trim();
        return s.Equals("T", StringComparison.OrdinalIgnoreCase) || s.Equals("Y", StringComparison.OrdinalIgnoreCase) ||
               s.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || s == "1";
    }
}
