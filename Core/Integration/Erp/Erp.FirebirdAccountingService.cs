namespace SKSB_App.Core.Integration.Erp;

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using FirebirdSql.Data.FirebirdClient;
using SKSB_App.Core.Diagnostics;

// =========================================================================
// 1. ACCOUNTING DTOs (read-only projection of SQL Accounting tables)
// =========================================================================

/// <summary>One row of SQL Accounting <c>PROJECT</c>.</summary>
public sealed record AccountingProjectRow(string Code, string Description, bool IsActive);

/// <summary>Offline-snapshot wrapper for active SQL Accounting project codes (avoids colliding with other <c>List&lt;string&gt;</c> snapshots).</summary>
public sealed record SqlAccountProjectCache(string Code);

/// <summary>
/// Read-only access to the SKSB SQL Accounting Firebird database (<c>ACC-0014.FDB</c>). Every query runs inside a
/// read-only, read-committed transaction; this service never writes to SQL Accounting.
/// </summary>
public class FirebirdAccountingService
{
    // =========================================================================
    // 2. EMBEDDED READ-ONLY CONNECTION CONFIG
    // =========================================================================
    public const string PrimaryHost = "NOR-HUAWEI";
    public const string FallbackHost = "192.168.0.222";
    public const int Port = 3050;
    public const string DatabasePath = @"C:\eStream\SQLAccounting\DB\ACC-0014.FDB";
    public const string UserId = "SKSB_RO";
    public const string Password = "s5PY8uhqJTYkv3lhaBV5";
    public const int ConnectionTimeoutSeconds = 5;
    public const int CommandTimeoutSeconds = 10;

    private const string SqlProjects = "SELECT CODE, DESCRIPTION, ISACTIVE FROM PROJECT ORDER BY CODE";

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
    // 3. PROJECT FETCH (primary host -> fallback host)
    // =========================================================================
    /// <summary>Every SQL Accounting project (blank codes skipped). Throws <see cref="InvalidOperationException"/> when both hosts fail.</summary>
    public List<AccountingProjectRow> FetchProjects()
    {
        return RunReadOnly("projects", (conn, trans) =>
        {
            var list = new List<AccountingProjectRow>();
            using var cmd = new FbCommand(SqlProjects, conn, trans) { CommandTimeout = CommandTimeoutSeconds };
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string code = Str(reader, 0);
                if (string.IsNullOrWhiteSpace(code)) continue;
                list.Add(new AccountingProjectRow(code, Str(reader, 1), Bool(reader, 2)));
            }
            return list;
        });
    }

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
                // Read-only transaction: nothing to commit, roll back to release the snapshot cleanly.
                try { trans.Rollback(); } catch { }
                AppLogger.Info($"[SQL Accounting] Read {label} from {host}.");
                return result;
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[SQL Accounting] Host '{host}' failed reading {label}: {ex.Message}");
                failures.Add($"{host}: {ex.Message}");
            }
        }
        throw new InvalidOperationException("SQL Accounting database is unreachable (" + string.Join(" | ", failures) + ").");
    }

    // =========================================================================
    // 4. READER HELPERS
    // =========================================================================
    private static string Str(IDataRecord r, int i) =>
        r.IsDBNull(i) ? string.Empty : (Convert.ToString(r.GetValue(i), CultureInfo.InvariantCulture) ?? string.Empty).Trim();

    /// <summary>SQL Accounting booleans are stored as 'T'/'F' chars (or 1/0 smallints, depending on version).</summary>
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
