using System.Windows.Forms;
using ExcelDna.Integration;

namespace SKSB_App.Core.Diagnostics
{
    public class LoadVerifier : IExcelAddIn
    {
        public void AutoOpen()
        {
            MessageBox.Show("Excel-DNA has successfully hooked into Excel process!\n\nxlAutoOpen executed.",
                            "Diagnostic: Add-In Loaded", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void AutoClose()
        {
        }
    }
}