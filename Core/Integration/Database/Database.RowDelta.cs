namespace SKSB_App.Core.Integration.Database;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Types;

// =========================================================================
// 1. CONCURRENCY CONFLICTS
// =========================================================================

/// <summary>
/// One column whose server value differs from the value this save would write. <see cref="ServerExcelValue"/> is
/// the server value in the shape written back to the sheet by "Keep Server Values" (null = empty cell).
/// </summary>
public sealed record FieldConflict(
    string ColumnHeader,
    string ServerDisplay,
    string LocalDisplay,
    object? ServerExcelValue = null);

/// <summary>
/// One row whose server copy changed (or appeared) since this workbook last loaded it. The Excel locator
/// (<see cref="ExcelTableName"/> / <see cref="KeyColumnHeader"/> / <see cref="RawKeyValue"/>) lets the save pipeline
/// point the Validation Task Pane at the exact conflicting cells.
/// </summary>
public sealed record ConcurrencyConflict(
    string TableName,
    string RecordKey,
    string Description,
    DateTime? ExpectedUpdatedAt,
    DateTime? CurrentServerUpdatedAt,
    string ExcelTableName = "",
    string KeyColumnHeader = "",
    string RawKeyValue = "",
    IReadOnlyList<FieldConflict>? FieldDiffs = null);

/// <summary>
/// Thrown inside a delta-save transaction when one or more optimistic-concurrency checks fail, so
/// <c>ExecuteInTransaction</c> rolls the whole batch back. Not a connection error: never triggers Offline Mode.
/// </summary>
public sealed class ConcurrencyConflictException : Exception
{
    public IReadOnlyList<ConcurrencyConflict> Conflicts { get; }

    public ConcurrencyConflictException(IReadOnlyList<ConcurrencyConflict> conflicts)
        : base(BuildMessage(conflicts))
    {
        Conflicts = conflicts;
    }

    private static string BuildMessage(IReadOnlyList<ConcurrencyConflict> conflicts)
    {
        if (conflicts.Count == 0) return "Concurrency conflict detected.";
        const int maxListed = 5;
        var lines = conflicts.Take(maxListed).Select(c => $"  - {c.Description}").ToList();
        if (conflicts.Count > maxListed) lines.Add($"  ... and {conflicts.Count - maxListed} more");
        return $"{conflicts.Count} record(s) were changed on the server by another user since this workbook was loaded:\n" +
               string.Join("\n", lines);
    }
}

// =========================================================================
// 2. TABLE DELTAS
// =========================================================================

/// <summary>A changed row plus the server <c>updated_at</c> it was based on (null = no concurrency token).</summary>
public sealed record ModifiedRow<TRecord>(TRecord Record, DateTime? ExpectedUpdatedAt);

public sealed class TableDelta<TRecord>
{
    /// <summary>False when no server baseline was captured (e.g. offline start): everything goes to <see cref="FallbackUpsert"/>.</summary>
    public bool HasBaseline { get; init; }
    public List<TRecord> Added { get; init; } = new();
    public List<ModifiedRow<TRecord>> Modified { get; init; } = new();
    public List<TRecord> FallbackUpsert { get; init; } = new();
    public int UnchangedCount { get; init; }

    public bool HasChanges => Added.Count > 0 || Modified.Count > 0 || FallbackUpsert.Count > 0;
}

/// <summary>
/// Snapshot of the rows (fingerprint + server <c>updated_at</c>) as last loaded from the server. Diffing the
/// current sheet rows against it yields the Added / Modified / Unchanged split used by delta saves.
/// </summary>
public sealed class TableBaseline<TKey> where TKey : notnull
{
    private readonly object _sync = new();
    private readonly Dictionary<TKey, (string Fingerprint, DateTime? UpdatedAt)> _rows = new();

    public bool HasBaseline { get; private set; }

    public void Capture<TRecord>(
        IEnumerable<TRecord> rows,
        Func<TRecord, TKey> keySelector,
        Func<TRecord, string> fingerprintSelector,
        Func<TRecord, DateTime?> updatedAtSelector)
    {
        var snapshot = new Dictionary<TKey, (string, DateTime?)>();
        foreach (var row in rows)
            snapshot[keySelector(row)] = (fingerprintSelector(row), updatedAtSelector(row));

        lock (_sync)
        {
            _rows.Clear();
            foreach (var kv in snapshot) _rows[kv.Key] = kv.Value;
            HasBaseline = true;
        }
    }

    public TableDelta<TRecord> Diff<TRecord>(
        IEnumerable<TRecord> currentRows,
        Func<TRecord, TKey> keySelector,
        Func<TRecord, string> fingerprintSelector)
    {
        lock (_sync)
        {
            if (!HasBaseline)
                return new TableDelta<TRecord> { HasBaseline = false, FallbackUpsert = currentRows.ToList() };

            var added = new List<TRecord>();
            var modified = new List<ModifiedRow<TRecord>>();
            int unchanged = 0;

            foreach (var row in currentRows)
            {
                if (!_rows.TryGetValue(keySelector(row), out var baseline))
                    added.Add(row);
                else if (!string.Equals(baseline.Fingerprint, fingerprintSelector(row), StringComparison.Ordinal))
                    modified.Add(new ModifiedRow<TRecord>(row, baseline.UpdatedAt));
                else
                    unchanged++;
            }

            return new TableDelta<TRecord>
            {
                HasBaseline = true,
                Added = added,
                Modified = modified,
                UnchangedCount = unchanged
            };
        }
    }

    /// <summary>Replaces (or adds) a single row's baseline, e.g. after the user accepted the server's copy of it.</summary>
    public void UpsertEntry(TKey key, string fingerprint, DateTime? updatedAt)
    {
        lock (_sync)
        {
            _rows[key] = (fingerprint, updatedAt);
            HasBaseline = true;
        }
    }

    public void Clear()
    {
        lock (_sync)
        {
            _rows.Clear();
            HasBaseline = false;
        }
    }
}

// =========================================================================
// 3. MODULE PAYLOADS
// =========================================================================

public sealed class StaffDeltaPayload
{
    public TableDelta<StaffRecord> Staff { get; init; } = new();
    public TableDelta<VehicleLicense> VehicleLicenses { get; init; } = new();
    public TableDelta<SkillLicense> SkillLicenses { get; init; } = new();
    public TableDelta<StaffSalary> Salaries { get; init; } = new();
    public TableDelta<StaffLeave> Leaves { get; init; } = new();
    public List<SkillLicenseType> SkillLicenseTypes { get; init; } = new();
}

public sealed class ProjectDeltaPayload
{
    public TableDelta<CompanyRecord> Companies { get; init; } = new();
    public TableDelta<ProjectRecord> Projects { get; init; } = new();
}

public sealed class VehicleDeltaPayload
{
    public TableDelta<VehicleRecord> Vehicles { get; init; } = new();
    public TableDelta<Maintenance> Maintenances { get; init; } = new();
    public TableDelta<RoadTax> RoadTaxes { get; init; } = new();
    public TableDelta<Inspection> Inspections { get; init; } = new();
    public TableDelta<VehiclePermit> Permits { get; init; } = new();
    public TableDelta<VehicleInsurance> Insurances { get; init; } = new();
    public TableDelta<InsuranceClaim> InsuranceClaims { get; init; } = new();
    public TableDelta<OdometerLog> OdometerLogs { get; init; } = new();
    public TableDelta<Workshop> Workshops { get; init; } = new();
    public TableDelta<InsuranceCompanyRecord> InsuranceCompanies { get; init; } = new();
    public TableDelta<FuelCardRecord> FuelCards { get; init; } = new();
    public List<string> FuelCompanies { get; init; } = new();
}

/// <summary>
/// Schedule delta: the four global reference tables are row deltas; the date-scoped board / settings / daily
/// attendance are written as one unit only when they changed (guarded by <see cref="ExpectedBoardMaxUpdatedAt"/>).
/// </summary>
public sealed class ScheduleDeltaPayload
{
    public ScheduleRecord Schedule { get; init; } = new();
    public bool HasDateBaseline { get; init; }
    public bool BoardOrAttendanceChanged { get; init; }
    public DateTime? ExpectedBoardMaxUpdatedAt { get; init; }
    public TableDelta<WorkerPreference> Preferences { get; init; } = new();
    public TableDelta<TeamRecord> Teams { get; init; } = new();
    public TableDelta<ScheduleSkillRecord> Skills { get; init; } = new();
    public TableDelta<ProjectSkillRequirement> ProjectSkillRequirements { get; init; } = new();
}
