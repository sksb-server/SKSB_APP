namespace SKSB_App.ProjectManager;

using System.Runtime.InteropServices;
using ExcelDna.Integration;
using SKSB_App.Core.RuntimeUI.Ribbon;
using SKSB_App.Core.WorkbookSave.Entry;

[ComVisible(true)]
public class RibbonController : ManagerRibbonBase
{
    // =========================================================================
    // 1. MODULE DESCRIPTOR
    // =========================================================================
    protected override ISaveContract Contract => Events.ProjectSyncDescriptor;
    protected override string TabId => "tabProjectManager";
    protected override string TabLabel => "Project Manager";

    protected override string PersistenceGroupXml => BuildPersistenceGroupXml(
        groupId: "grpProjectPersistence",
        groupLabel: "Project Database Sync",
        loadId: "btnLoadProjects",
        loadScreentip: "Load Project Records",
        loadSupertip: "Fetches project masters, daily summaries and companies from PostgreSQL or offline storage.",
        saveId: "btnSaveProjects",
        saveScreentip: "Save Project Records",
        saveSupertip: "Validates and commits all project sheets, daily summaries and companies.");

    // An empty fetch still populates (clears) the worksheets and reloads the Companies table.
    protected override bool PopulateWhenEmpty => true;
    protected override bool ShowEmptyLoadDialog => false;

    // =========================================================================
    // 2. EXCEL COMMANDS (must live on a non-generic type)
    // =========================================================================
    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro() => Events.RunDelayedAutoLoad();

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave() => Events.RunHotKeySave();
}
