namespace SKSB_App.Core.WorkbookOpen.Entry;

using System.Windows.Forms;
using ExcelDna.Integration;

#if DEBUG
public class DiagnosticsVerifier : IExcelAddIn
{
    public void AutoOpen()
    {
        MessageBox.Show(
            "Excel-DNA has successfully hooked into Excel process!\n\nxlAutoOpen executed.",
            "Diagnostic: Add-In Loaded",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    public void AutoClose()
    {
    }
}
#endif