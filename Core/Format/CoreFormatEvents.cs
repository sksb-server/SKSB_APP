namespace SKSB_App.Core;

using System;
using ExcelDna.Integration;
using SKSB_App.Core.UI;

public class CoreEvents : IExcelAddIn
{
    private dynamic _xlApp;

    public void AutoOpen()
    {
        _xlApp = ExcelDnaUtil.Application;

        // Attach the double-click event listener using dynamic COM interop
        _xlApp.SheetBeforeDoubleClick += new COMDoubleClickDelegate(OnSheetBeforeDoubleClick);
    }

    public void AutoClose()
    {
        if (_xlApp != null)
        {
            _xlApp.SheetBeforeDoubleClick -= new COMDoubleClickDelegate(OnSheetBeforeDoubleClick);
            _xlApp = null;
        }
    }

    // Delegate signature required for Excel's SheetBeforeDoubleClick
    public delegate void COMDoubleClickDelegate(dynamic Sh, dynamic Target, ref bool Cancel);

    private void OnSheetBeforeDoubleClick(dynamic Sh, dynamic Target, ref bool Cancel)
    {
        try
        {
            // Check if the cell has data validation
            if (Target.Validation != null && Target.Validation.Type == 3) // 3 = xlValidateList
            { 
                // We assume it's our custom multi-select if the error alert is disabled
                if (Target.Validation.ShowError == false)
                {
                    // Cancel native cell editing
                    Cancel = true;

                    // Show our WinForms multi-select interface
                    FormatCommands.ShowChecklistForActiveCell(Target);
                }
            }
        }
        catch
        {
            // Ignore cells without validation (Target.Validation throws an exception if none exists)
        }
    }
}