namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Erp;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.Types;
using SKSB_App.Core.Types.Attributes;

// =========================================================================
// 1. SQL PAYROLL MAPPING RULES
// =========================================================================
public static class StaffImportConfig
{
    public const string DefaultDepartment = "Worker";

    /// <summary>Daily-rated salaries never exceed this; larger amounts are monthly salaries keyed with the wrong type.</summary>
    public const decimal DailySalaryAnomalyThreshold = 500m;

    /// <summary>Thumbprint controller name field length.</summary>
    public const int TerminalNameMaxLength = 23;

    public const string HrGroupOfficeStaff = "OFFICE STAFF";
    public const string HrGroupSiteWorker = "SITE WORKER";

    private static readonly string[] AllowedDepartments = { "Worker", "Admin", "Office" };

    /// <summary>Exact 1:1 mapping (case-insensitive) of SQL Payroll Department to Worker / Admin / Office; null otherwise.</summary>
    public static string? MapDepartment(string? rawDepartment)
    {
        if (string.IsNullOrWhiteSpace(rawDepartment)) return null;
        string clean = rawDepartment.Trim();
        return AllowedDepartments.FirstOrDefault(d => d.Equals(clean, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Parses SQL Payroll WAGESTYPE into a <see cref="SalaryBasis"/>. Returns null when the type is not recognised.</summary>
    public static SalaryBasis? ParseSalaryType(string? rawType)
    {
        string t = (rawType ?? string.Empty).Trim().ToUpperInvariant();
        return t switch
        {
            "DW-X" or "DAILY" => SalaryBasis.Daily,
            "MONTHLY" or "B.WAGES" or "DR" => SalaryBasis.Monthly,
            _ => null
        };
    }

    /// <summary>Basis for an amount: Daily amounts above the anomaly threshold are treated as Monthly.</summary>
    public static SalaryBasis ResolveBasis(string? rawType, decimal amount, out bool recognised, out bool anomaly)
    {
        SalaryBasis? parsed = ParseSalaryType(rawType);
        recognised = parsed.HasValue;
        SalaryBasis basis = parsed ?? SalaryBasis.Monthly;
        anomaly = basis == SalaryBasis.Daily && amount > DailySalaryAnomalyThreshold;
        return anomaly ? SalaryBasis.Monthly : basis;
    }
}

// =========================================================================
// 2. SYNC RESULT
// =========================================================================
public class StaffImportResult
{
    public List<StaffRecord> ImportedStaffs { get; set; } = new();
    public List<ImportSkippedRecord> AdvisoryItems { get; set; } = new();
    public bool PayrollConnected { get; set; }
    public bool ThumbprintChecked { get; set; }
    public string SourceDescription { get; set; } = string.Empty;

    public const string AdvisoryBannerText =
        "Please fix the issues below directly in SQL Payroll (or on the thumbprint controller). " +
        "SKSB_App mirrors SQL Payroll read-only: until they are corrected there, the staff data viewed here " +
        "and committed to the system will be untrue or incomplete.";

    public ImportSummaryReport ToGenericReport() => new()
    {
        SourceSystemName = "SQL Payroll",
        EntityLabel = Loc.T("staff.import.entityLabel"),
        SuccessfullyImportedCount = ImportedStaffs.Count,
        FilteredExclusionCount = 0,
        FilterRuleDescription = Loc.T("staff.import.filterRule"),
        SkippedErrors = AdvisoryItems,
        AdvisoryBanner = Loc.T("staff.import.advisoryBanner"),
        SkippedSectionTitle = Loc.T("staff.import.skippedSectionTitle"),
        Title = AdvisoryItems.Count > 0 ? Loc.T("staff.import.titleAdvisories") : Loc.T("staff.import.titleSuccess"),
        FooterHint = Loc.T("staff.import.footerHint")
    };
}

// =========================================================================
// 3. SQL PAYROLL -> STAFF RECORDS (read-only direct Firebird sync)
// =========================================================================
public static class Importer
{
    private static readonly Regex RxReviseWages = new(@"Revise Wages from\s+([\d,\.]+)\s+to\s+([\d,\.]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RxReviseWagesType = new(@"Revise Wages Type from\s+(\S+)\s+to\s+(\S+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private const string CatStatus = "Status";
    private const string CatJoinDate = "Join Date";
    private const string CatDepartment = "Department";
    private const string CatSalaries = "Salaries";
    private const string CatProfile = "Profile";
    private const string CatLeaves = "Leaves";
    private const string CatThumbprint = "Thumbprint";

    /// <summary>
    /// Reads SQL Payroll directly (read-only) and rebuilds every staff profile + salary history, keeping licenses and
    /// server IDs / concurrency tokens from <paramref name="existingPgStaffs"/>. Throws when SQL Payroll is unreachable.
    /// </summary>
    public static StaffImportResult SyncFromSqlPayroll(List<StaffRecord>? existingPgStaffs)
    {
        var result = new StaffImportResult();
        var today = DateOnly.FromDateTime(DateTime.Today);

        var snapshot = new FirebirdPayrollService().FetchPayrollSnapshot();
        result.PayrollConnected = true;
        result.SourceDescription = $"SQL Payroll (PAY-0003.FDB @ {snapshot.ConnectedHost})";

        var pgById = (existingPgStaffs ?? new List<StaffRecord>())
            .Where(s => s.StaffID > 0)
            .GroupBy(s => s.StaffID)
            .ToDictionary(g => g.Key, g => g.First());

        var historyByCode = snapshot.WageHistory
            .GroupBy(h => (h.Code ?? string.Empty).Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.OrderBy(h => h.PostDate).ThenBy(h => h.AutoKey).ToList(), StringComparer.OrdinalIgnoreCase);

        var leavesByCode = snapshot.Leaves
            .GroupBy(l => NormalizeEnroll(l.Code), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        var seen = new HashSet<long>();

        foreach (var emp in snapshot.Employees)
        {
            string code = (emp.Code ?? string.Empty).Trim();
            string name = (emp.Name ?? string.Empty).Trim();

            // 3a. Identity
            if (!long.TryParse(code, NumberStyles.Integer, CultureInfo.InvariantCulture, out long staffId) || staffId <= 0)
            {
                Advise(result, CatProfile, code.Length == 0 ? "[Empty]" : code, name,
                    $"Employee Code '{code}' is not a positive number, so it cannot be used as a Staff ID. Skipped. Change the Code in SQL Payroll to a numeric Staff ID.");
                continue;
            }
            if (!seen.Add(staffId))
            {
                Advise(result, CatProfile, code, name, $"Duplicate Staff ID {staffId} in SQL Payroll (codes with leading zeros?). Only the first row was used.");
                continue;
            }
            if (name.Length == 0)
            {
                Advise(result, CatProfile, code, "[Empty]", $"Staff ID {staffId} has no Name in SQL Payroll. Skipped.");
                continue;
            }

            pgById.TryGetValue(staffId, out var pg);

            // 3b. Employment status (strictly driven by Resign Date)
            //     Inactive with no Resign Date is auto-evaluated as Suspended.
            ContractStatus status = emp.ResignDate.HasValue && emp.ResignDate.Value <= today
                ? ContractStatus.Resigned
                : (!emp.IsActive && !emp.ResignDate.HasValue ? ContractStatus.Suspended : ContractStatus.Active);

            if (emp.IsActive && emp.ResignDate.HasValue)
            {
                Advise(result, CatStatus, code, name,
                    $"Marked Active in SQL Payroll but has a Resign Date ({emp.ResignDate.Value:dd/MM/yyyy}). SKSB_App follows the Resign Date " +
                    $"(treated as {(status == ContractStatus.Resigned ? "Resigned" : "Active until that date")}). HR: untick 'Active' in SQL Payroll.");
            }
            else if (!emp.IsActive && !emp.ResignDate.HasValue)
            {
                Advise(result, CatStatus, code, name,
                    "Marked Inactive in SQL Payroll with no Resign Date — auto-evaluated as Suspended. " +
                    "HR: if resigned, set Resign Date in SQL Payroll; if still active, tick 'Active' in SQL Payroll.");
            }

            DateOnly joinDate;
            if (emp.JoinDate.HasValue)
            {
                joinDate = emp.JoinDate.Value;
            }
            else
            {
                joinDate = pg?.StaffDetails?.JoinDate ?? today;
                if (emp.ResignDate.HasValue && emp.ResignDate.Value < joinDate) joinDate = emp.ResignDate.Value;
                if (status == ContractStatus.Active)
                {
                    Advise(result, CatJoinDate, code, name,
                        $"Active staff has no Join Date in SQL Payroll (using {joinDate:dd/MM/yyyy}). HR: set the Join Date in SQL Payroll so salary history and attendance resolve correctly.");
                }
            }
            if (emp.ResignDate.HasValue && emp.ResignDate.Value < joinDate)
            {
                Advise(result, CatJoinDate, code, name,
                    $"Resign Date ({emp.ResignDate.Value:dd/MM/yyyy}) is earlier than Join Date ({joinDate:dd/MM/yyyy}) in SQL Payroll. Join Date was set to the Resign Date. HR: correct the dates in SQL Payroll.");
                joinDate = emp.ResignDate.Value;
            }

            // 3c. Department (exact Worker / Admin / Office)
            string? mappedDept = StaffImportConfig.MapDepartment(emp.Department);
            string department = mappedDept ?? StaffImportConfig.DefaultDepartment;
            string rawDept = string.IsNullOrWhiteSpace(emp.Department) ? "(blank)" : emp.Department.Trim();
            string hrGroup = (emp.HrGroup ?? string.Empty).Trim();

            if (mappedDept == null)
            {
                Advise(result, CatDepartment, code, name,
                    $"Department '{rawDept}' is not recognised (shown as '{StaffImportConfig.DefaultDepartment}'). HR: change the Department in SQL Payroll to Worker, Admin, or Office.");
            }

            bool deptIsWorkerLike = string.Equals(mappedDept, "Worker", StringComparison.OrdinalIgnoreCase) ||
                                    rawDept.Equals("PROJECT", StringComparison.OrdinalIgnoreCase);
            bool deptIsOfficeLike = string.Equals(mappedDept, "Office", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(mappedDept, "Admin", StringComparison.OrdinalIgnoreCase);
            if (deptIsWorkerLike && hrGroup.Equals(StaffImportConfig.HrGroupOfficeStaff, StringComparison.OrdinalIgnoreCase))
            {
                Advise(result, CatDepartment, code, name,
                    $"Department '{rawDept}' conflicts with HR Group '{hrGroup}'. HR: set Department to Office/Admin, or correct the HR Group in SQL Payroll.");
            }
            else if (deptIsOfficeLike && hrGroup.Equals(StaffImportConfig.HrGroupSiteWorker, StringComparison.OrdinalIgnoreCase))
            {
                Advise(result, CatDepartment, code, name,
                    $"Department '{rawDept}' conflicts with HR Group '{hrGroup}'. HR: set Department to Worker, or correct the HR Group in SQL Payroll.");
            }

            // 3d. Contact & identity
            long nric = 0;
            if (!string.IsNullOrWhiteSpace(emp.NewIc))
            {
                string cleanIc = Regex.Replace(emp.NewIc, @"[^\d]", string.Empty);
                long.TryParse(cleanIc, NumberStyles.Integer, CultureInfo.InvariantCulture, out nric);
            }
            string phone = !string.IsNullOrWhiteSpace(emp.Mobile) ? emp.Mobile.Trim() : (emp.Phone1 ?? string.Empty).Trim();

            var staff = new StaffRecord
            {
                StaffID = staffId,
                StaffDetails = new StaffDetails
                {
                    StaffName = name,
                    StaffShortName = !string.IsNullOrWhiteSpace(pg?.StaffDetails?.StaffShortName)
                        ? pg!.StaffDetails.StaffShortName
                        : RuleEvaluator.ApplyTextTransform(name, TransformKind.ShortNameFirstLast),
                    Department = department,
                    NRIC = nric,
                    Passport = (emp.Passport ?? string.Empty).Trim(),
                    Phone = phone,
                    JoinDate = joinDate,
                    ResignDate = emp.ResignDate,
                    ContractStatus = status
                },
                VehicleLicenses = pg?.VehicleLicenses ?? new List<VehicleLicense>(),
                SkillLicenses = pg?.SkillLicenses ?? new List<SkillLicense>(),
                ServerUpdatedAt = pg?.ServerUpdatedAt
            };

            // 3e. Salaries + reconstructed history
            historyByCode.TryGetValue(code, out var history);
            staff.Salaries = BuildSalaryHistory(result, emp, staff, history ?? new List<PayrollWageHistoryRow>(), pg, today);
            staff.RefreshCurrentSalary();

            // 3f. Leaves (LV_APP), keeping server updated_at tokens by payroll_dockey.
            leavesByCode.TryGetValue(staffId.ToString(CultureInfo.InvariantCulture), out var empLeaves);
            staff.Leaves = BuildLeaves(empLeaves, staff, pg);
            foreach (var lv in staff.Leaves.Where(l => l.LeaveDay <= 0))
            {
                Advise(result, CatLeaves, staffId.ToString(CultureInfo.InvariantCulture), staff.StaffDetails.StaffName,
                    $"Legacy SQL Payroll leave entry on {lv.LeaveDate:yyyy-MM-dd} (DOCKEY {lv.LeaveID}) has negative/zero Days ({lv.LeaveDay}). Imported as-is.");
            }

            result.ImportedStaffs.Add(staff);
        }

        // 3g. Keep SKSB_App-only staff (no hard deletes) but flag them.
        foreach (var pg in pgById.Values.Where(p => !seen.Contains(p.StaffID)).OrderBy(p => p.StaffID))
        {
            pg.RefreshCurrentSalary();
            result.ImportedStaffs.Add(pg);
            if (pg.StaffDetails.ContractStatus != ContractStatus.Resigned)
            {
                Advise(result, CatProfile, pg.StaffID.ToString(CultureInfo.InvariantCulture), pg.StaffDetails.StaffName,
                    "Exists in SKSB_App but not in SQL Payroll (kept as-is). HR: add this employee in SQL Payroll, or set a Resign Date.");
            }
        }

        CheckThumbprintTerminal(result);
        result.ImportedStaffs = result.ImportedStaffs.OrderBy(s => s.StaffID).ToList();

        AppLogger.Info($"[StaffManager] SQL Payroll sync: {result.ImportedStaffs.Count} staff, " +
            $"{result.ImportedStaffs.Sum(s => s.Salaries.Count)} salary row(s), {result.ImportedStaffs.Sum(s => s.Leaves.Count)} leave row(s), {result.AdvisoryItems.Count} advisory item(s), thumbprint checked={result.ThumbprintChecked}.");
        return result;
    }

    // =========================================================================
    // 4. SALARY HISTORY RECONSTRUCTION (HR_EMP + HR_EMP_HISTORY)
    // =========================================================================
    private sealed record SalaryEvent(DateOnly Date, long AutoKey, decimal Amount, SalaryBasis Basis, string Remarks);

    private static List<StaffSalary> BuildSalaryHistory(
        StaffImportResult result, PayrollEmployeeRow emp, StaffRecord staff, List<PayrollWageHistoryRow> history, StaffRecord? pg, DateOnly today)
    {
        string code = emp.Code.Trim();
        string name = staff.StaffDetails.StaffName;
        DateOnly joinDate = staff.StaffDetails.JoinDate;

        // Current salary (HR_EMP.WAGES / WAGESTYPE)
        SalaryBasis currentBasis = StaffImportConfig.ResolveBasis(emp.WagesType, emp.Wages, out bool recognised, out bool anomaly);
        if (!recognised)
        {
            Advise(result, CatSalaries, code, name,
                $"Wages Type '{(string.IsNullOrWhiteSpace(emp.WagesType) ? "(blank)" : emp.WagesType)}' is not recognised (treated as Monthly). HR: set a valid Wages Type (DAILY, DW-X, MONTHLY, B.WAGES, DR) in SQL Payroll.");
        }
        if (anomaly)
        {
            Advise(result, CatSalaries, code, name,
                $"Wages Type '{emp.WagesType}' is daily but the salary is RM {emp.Wages:#,##0.00} (looks like a monthly salary; treated as Monthly). HR: fix the Wages Type in SQL Payroll.");
        }
        if (staff.StaffDetails.ContractStatus == ContractStatus.Active && emp.Wages <= 0m)
        {
            Advise(result, CatSalaries, code, name, "Active staff has no salary (RM 0.00) in SQL Payroll. HR: set the Wages in SQL Payroll so salary can be calculated.");
        }

        // Walk history chronologically, tracking the salary type in force at each revision.
        var typeRevisions = history
            .Select(h => (Row: h, Match: RxReviseWagesType.Match(h.Description ?? string.Empty)))
            .Where(x => x.Match.Success)
            .ToList();
        string typeInForce = typeRevisions.Count > 0 ? typeRevisions[0].Match.Groups[1].Value : emp.WagesType;
        string initialType = typeInForce;

        var revisions = new List<SalaryEvent>();
        decimal? firstFromAmount = null;
        foreach (var h in history)
        {
            string desc = h.Description ?? string.Empty;
            var mType = RxReviseWagesType.Match(desc);
            if (mType.Success) typeInForce = mType.Groups[2].Value;

            var mSalary = RxReviseWages.Match(desc);
            if (!mSalary.Success) continue;
            if (!TryParseAmount(mSalary.Groups[1].Value, out decimal fromAmt) || !TryParseAmount(mSalary.Groups[2].Value, out decimal toAmt)) continue;

            firstFromAmount ??= fromAmt;
            var basis = StaffImportConfig.ResolveBasis(typeInForce, toAmt, out _, out _);
            revisions.Add(new SalaryEvent(h.PostDate, h.AutoKey, toAmt, basis, "Revised in SQL Payroll"));
        }

        // Same-date revisions: keep the latest AutoKey.
        revisions = revisions
            .GroupBy(r => r.Date)
            .Select(g => g.OrderByDescending(r => r.AutoKey).First())
            .OrderBy(r => r.Date)
            .ToList();

        var events = new List<SalaryEvent>();
        if (revisions.Count > 0)
        {
            if (firstFromAmount.HasValue && firstFromAmount.Value > 0m && joinDate < revisions[0].Date)
            {
                var basis = StaffImportConfig.ResolveBasis(initialType, firstFromAmount.Value, out _, out _);
                events.Add(new SalaryEvent(joinDate, 0, firstFromAmount.Value, basis, "Initial (SQL Payroll history)"));
            }
            events.AddRange(revisions);

            var last = events[^1];
            if (emp.Wages > 0m && (last.Amount != emp.Wages || last.Basis != currentBasis))
            {
                if (last.Date >= today) events.RemoveAt(events.Count - 1);
                events.Add(new SalaryEvent(today, long.MaxValue, emp.Wages, currentBasis, "Current (SQL Payroll)"));
            }
        }
        else if (emp.Wages > 0m)
        {
            events.Add(new SalaryEvent(joinDate, 0, emp.Wages, currentBasis, "SQL Payroll"));
        }

        // Map to StaffSalary rows, preserving the PostgreSQL salary_id / updated_at of identical rows.
        var pgSalaries = (pg?.Salaries ?? new List<StaffSalary>())
            .GroupBy(s => SalaryKey(s.EffectiveDate, s.SalaryBasis, s.SalaryRate))
            .ToDictionary(g => g.Key, g => g.First());

        // Chronological walk: each row's "Salary Before" is the immediately preceding row (null for the initial row).
        var rows = new List<StaffSalary>();
        StaffSalary? previous = null;
        foreach (var e in events.Where(e => e.Amount > 0m))
        {
            decimal rate = decimal.Round(e.Amount, 2, MidpointRounding.AwayFromZero);
            pgSalaries.TryGetValue(SalaryKey(e.Date, e.Basis, rate), out var existing);
            var row = new StaffSalary
            {
                SalaryID = existing?.SalaryID ?? 0,
                StaffID = staff.StaffID,
                StaffName = name,
                EffectiveDate = e.Date,
                SalaryBefore = previous?.SalaryRate,
                SalaryBeforeBasis = previous?.SalaryBasis ?? SalaryBasis.None,
                SalaryBasis = e.Basis,
                SalaryRate = rate,
                Remarks = e.Remarks,
                ServerUpdatedAt = existing?.ServerUpdatedAt
            };
            rows.Add(row);
            previous = row;
        }
        return rows;
    }

    /// <summary>Maps LV_APP rows to <see cref="StaffLeave"/> (newest first), preserving server <c>updated_at</c> by Leave ID.</summary>
    private static List<StaffLeave> BuildLeaves(List<PayrollLeaveRow>? rows, StaffRecord staff, StaffRecord? pg)
    {
        if (rows == null || rows.Count == 0) return new List<StaffLeave>();

        var pgLeaves = (pg?.Leaves ?? new List<StaffLeave>())
            .Where(l => l.LeaveID > 0)
            .GroupBy(l => l.LeaveID)
            .ToDictionary(g => g.Key, g => g.First());

        return rows
            .Where(r => r.DocKey > 0)
            .GroupBy(r => r.DocKey)
            .Select(g => g.First())
            .Select(r => new StaffLeave
            {
                LeaveID = r.DocKey,
                StaffID = staff.StaffID,
                StaffName = staff.StaffDetails.StaffName,
                LeaveDate = r.LeaveDate,
                LeaveType = PayrollLeaveMapper.MapStatus(r.LeaveType).ToString(),
                LeaveDay = r.LeaveDay,
                Description = (r.Description ?? string.Empty).Trim(),
                ServerUpdatedAt = pgLeaves.TryGetValue(r.DocKey, out var existing) ? existing.ServerUpdatedAt : null
            })
            .OrderByDescending(l => l.LeaveDate)
            .ThenByDescending(l => l.LeaveID)
            .ToList();
    }

    private static string SalaryKey(DateOnly date, SalaryBasis basis, decimal rate) =>
        $"{date:yyyy-MM-dd}|{basis}|{decimal.Round(rate, 2, MidpointRounding.AwayFromZero).ToString("F2", CultureInfo.InvariantCulture)}";

    private static bool TryParseAmount(string raw, out decimal amount) =>
        decimal.TryParse((raw ?? string.Empty).Replace(",", string.Empty).TrimEnd('.'),
            NumberStyles.Number, CultureInfo.InvariantCulture, out amount);

    // =========================================================================
    // 5. THUMBPRINT CONTROLLER MISMATCH CHECK
    // =========================================================================
    private static void CheckThumbprintTerminal(StaffImportResult result)
    {
        List<TerminalUser> terminalUsers;
        try
        {
            terminalUsers = DeviceSyncManager.FetchTerminalUsersAsync().GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[StaffManager] Thumbprint controller check skipped: {ex.Message}");
            Advise(result, CatThumbprint, DeviceSyncManager.DefaultIp, "Thumbprint Controller",
                $"The thumbprint controller ({DeviceSyncManager.DefaultIp}) could not be reached, so staff IDs/names on it were not checked. ({ex.Message})");
            return;
        }

        result.ThumbprintChecked = true;
        var staffById = result.ImportedStaffs
            .GroupBy(s => s.StaffID.ToString(CultureInfo.InvariantCulture))
            .ToDictionary(g => g.Key, g => g.First());
        var terminalIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int nameMismatches = 0;

        foreach (var u in terminalUsers)
        {
            string enroll = NormalizeEnroll(u.EnrollNumber);
            if (enroll.Length == 0) continue;
            terminalIds.Add(enroll);
            string terminalName = (u.Name ?? string.Empty).Trim();

            if (!staffById.TryGetValue(enroll, out var staff))
            {
                Advise(result, CatThumbprint, enroll, terminalName,
                    $"Thumbprint user ID {enroll} ('{terminalName}') does not match any Staff ID in SQL Payroll. Fix the ID on the controller or add the employee in SQL Payroll.");
                continue;
            }

            var contractStatus = staff.StaffDetails.ContractStatus;
            if (contractStatus != ContractStatus.Active)
            {
                string label = contractStatus == ContractStatus.Suspended ? "Suspended" : "Resigned";
                Advise(result, CatThumbprint, enroll, staff.StaffDetails.StaffName,
                    $"{label} staff is still enrolled on the thumbprint controller. Remove the user from the controller.");
                continue;
            }

            string expected = ExpectedTerminalName(staff);
            if (expected.Length > 0 && !string.Equals(terminalName, expected, StringComparison.Ordinal))
            {
                nameMismatches++;
                Advise(result, CatThumbprint, enroll, staff.StaffDetails.StaffName,
                    $"Controller name '{terminalName}' differs from Short Name '{expected}'. Use 'Push ID & Name to Controller' to update it.");
            }
        }

        foreach (var s in result.ImportedStaffs.Where(s =>
                     s.StaffDetails.ContractStatus == ContractStatus.Active &&
                     string.Equals(s.StaffDetails.Department, "Worker", StringComparison.OrdinalIgnoreCase)))
        {
            string id = s.StaffID.ToString(CultureInfo.InvariantCulture);
            if (!terminalIds.Contains(id))
            {
                Advise(result, CatThumbprint, id, s.StaffDetails.StaffName,
                    "Active worker is not enrolled on the thumbprint controller, so attendance cannot be recorded. Enrol the worker on the controller.");
            }
        }

        AppLogger.Info($"[StaffManager] Thumbprint check: {terminalUsers.Count} controller user(s), {nameMismatches} name mismatch(es).");
    }

    private static string NormalizeEnroll(string? enroll)
    {
        string e = (enroll ?? string.Empty).Trim();
        return long.TryParse(e, NumberStyles.Integer, CultureInfo.InvariantCulture, out long n)
            ? n.ToString(CultureInfo.InvariantCulture)
            : e;
    }

    private static string ExpectedTerminalName(StaffRecord s)
    {
        string name = !string.IsNullOrWhiteSpace(s.StaffDetails.StaffShortName)
            ? s.StaffDetails.StaffShortName.Trim()
            : (s.StaffDetails.StaffName ?? string.Empty).Trim();
        return name.Length > StaffImportConfig.TerminalNameMaxLength ? name.Substring(0, StaffImportConfig.TerminalNameMaxLength) : name;
    }

    // =========================================================================
    // 6. ADVISORY HELPER
    // =========================================================================
    private static void Advise(StaffImportResult result, string category, string identifier, string name, string reason) =>
        result.AdvisoryItems.Add(new ImportSkippedRecord
        {
            ExcelRowNumber = 0,
            Category = category,
            Identifier = identifier,
            NameOrTitle = name,
            Reason = reason
        });
}
