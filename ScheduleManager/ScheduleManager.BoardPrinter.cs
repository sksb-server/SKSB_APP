namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;

internal static class BoardPrinter
{
    // =========================================================================
    // 1. PDF EXPORT
    // =========================================================================
    internal static void ExportPrintableSchedulePdf()
    {
        string? pdfPath = null;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            object? wbObj = xlApp.ActiveWorkbook;
            object? wsObj = wbObj == null ? null : (object?)SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            if (wbObj == null || wsObj == null)
            {
                MessageBox.Show("The 'Daily Schedule' sheet was not found in the active workbook.", "Print Schedule", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            dynamic wb = wbObj;
            dynamic ws = wsObj;

            DateOnly date = ScheduleBoardMapper.ReadDate((object)ws) ?? DateOnly.FromDateTime(DateTime.Today);
            var slots = ScheduleBoardMapper.Read((object)ws, date);

            bool dayHasData = slots.Any(s => s.Shift == Shifts.Morning);
            bool nightHasData = slots.Any(s => s.Shift == Shifts.Night);
            int activeSlotCount = slots.Count > 0 ? slots.Max(s => s.SlotIndex) + 1 : 1;
            int lastCol = ScheduleBoardMapper.SlotColumn(activeSlotCount - 1) + 1;
            int lastRow = nightHasData ? 34 : 18;
            bool hideDayRows = !dayHasData && nightHasData;

            Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "SKSB_Print"));
            pdfPath = Path.Combine(Path.GetTempPath(), "SKSB_Print", $"Daily_Schedule_{date:yyyy-MM-dd}_{DateTime.Now:HHmmss}.pdf");

            Arrow.HideAll();

            // Capture Excel state
            bool enableEvents = true, screenUpdating = true, displayAlerts = true, printComm = true, wasSaved = false;
            int calc = -4105;
            try { enableEvents = (bool)xlApp.EnableEvents; } catch { }
            try { screenUpdating = (bool)xlApp.ScreenUpdating; } catch { }
            try { calc = (int)xlApp.Calculation; } catch { }
            try { displayAlerts = (bool)xlApp.DisplayAlerts; } catch { }
            try { printComm = (bool)xlApp.PrintCommunication; } catch { }
            try { wasSaved = (bool)wb.Saved; } catch { }

            var hiddenCols = new List<(int Col, bool WasHidden)>();
            bool rowsHidden = false, rowsWereHidden = false;
            dynamic ps = null!;
            bool psSet = false;
            object? oPrintArea = null, oOrientation = null, oPaper = null, oZoom = null, oWide = null, oTall = null,
                oCenterH = null, oCenterV = null, oComments = null, oL = null, oR = null, oT = null, oB = null, oHdr = null, oFtr = null;

            try
            {
                try { xlApp.EnableEvents = false; } catch { }
                try { xlApp.ScreenUpdating = false; } catch { }
                try { xlApp.Calculation = -4135; } catch { }
                try { xlApp.DisplayAlerts = false; } catch { }

                for (int slot = 0; slot < activeSlotCount; slot++)
                {
                    int col = ScheduleBoardMapper.SlotColumn(slot) + 2;
                    bool was = false;
                    try { was = (bool)ws.Columns[col].Hidden; } catch { }
                    hiddenCols.Add((col, was));
                    try { ws.Columns[col].Hidden = true; } catch { }
                }

                if (hideDayRows)
                {
                    try { rowsWereHidden = (bool)ws.Rows["4:19"].Hidden; } catch { }
                    try { ws.Rows["4:19"].Hidden = true; rowsHidden = true; } catch { }
                }

                ps = ws.PageSetup;
                psSet = true;
                try { xlApp.PrintCommunication = false; } catch { }
                try { oPrintArea = ps.PrintArea; } catch { }
                try { oOrientation = ps.Orientation; } catch { }
                try { oPaper = ps.PaperSize; } catch { }
                try { oZoom = ps.Zoom; } catch { }
                try { oWide = ps.FitToPagesWide; } catch { }
                try { oTall = ps.FitToPagesTall; } catch { }
                try { oCenterH = ps.CenterHorizontally; } catch { }
                try { oCenterV = ps.CenterVertically; } catch { }
                try { oComments = ps.PrintComments; } catch { }
                try { oL = ps.LeftMargin; } catch { }
                try { oR = ps.RightMargin; } catch { }
                try { oT = ps.TopMargin; } catch { }
                try { oB = ps.BottomMargin; } catch { }
                try { oHdr = ps.HeaderMargin; } catch { }
                try { oFtr = ps.FooterMargin; } catch { }

                ps.PrintArea = $"$A$1:${ScheduleBoardMapper.ColumnLetter(lastCol)}${lastRow}";
                ps.Orientation = 2;
                ps.PaperSize = 9;
                ps.Zoom = false;
                ps.FitToPagesWide = 1;
                ps.FitToPagesTall = 1;
                ps.CenterHorizontally = true;
                ps.CenterVertically = false;
                ps.PrintComments = -4142;
                ps.LeftMargin = xlApp.InchesToPoints(0.3);
                ps.RightMargin = xlApp.InchesToPoints(0.3);
                ps.TopMargin = xlApp.InchesToPoints(0.35);
                ps.BottomMargin = xlApp.InchesToPoints(0.35);
                ps.HeaderMargin = xlApp.InchesToPoints(0.2);
                ps.FooterMargin = xlApp.InchesToPoints(0.2);
                try { xlApp.PrintCommunication = true; } catch { }

                ws.ExportAsFixedFormat(0, pdfPath, 0, true, false, Type.Missing, Type.Missing, false);
            }
            finally
            {
                foreach (var (col, was) in hiddenCols)
                    try { ws.Columns[col].Hidden = was; } catch { }
                if (rowsHidden)
                    try { ws.Rows["4:19"].Hidden = rowsWereHidden; } catch { }

                if (psSet)
                {
                    try { xlApp.PrintCommunication = false; } catch { }
                    try { if (oPrintArea is string pa) ps.PrintArea = pa; else ps.PrintArea = false; } catch { }
                    try { if (oOrientation != null) ps.Orientation = oOrientation; } catch { }
                    try { if (oPaper != null) ps.PaperSize = oPaper; } catch { }
                    try
                    {
                        if (oZoom is bool zb && !zb)
                        {
                            ps.Zoom = false;
                            if (oWide != null) ps.FitToPagesWide = oWide;
                            if (oTall != null) ps.FitToPagesTall = oTall;
                        }
                        else if (oZoom != null) ps.Zoom = oZoom;
                    }
                    catch { }
                    try { if (oCenterH != null) ps.CenterHorizontally = oCenterH; } catch { }
                    try { if (oCenterV != null) ps.CenterVertically = oCenterV; } catch { }
                    try { if (oComments != null) ps.PrintComments = oComments; } catch { }
                    try { if (oL != null) ps.LeftMargin = oL; } catch { }
                    try { if (oR != null) ps.RightMargin = oR; } catch { }
                    try { if (oT != null) ps.TopMargin = oT; } catch { }
                    try { if (oB != null) ps.BottomMargin = oB; } catch { }
                    try { if (oHdr != null) ps.HeaderMargin = oHdr; } catch { }
                    try { if (oFtr != null) ps.FooterMargin = oFtr; } catch { }
                }
                try { xlApp.PrintCommunication = printComm; } catch { }
                try { xlApp.Calculation = calc; } catch { }
                try { xlApp.DisplayAlerts = displayAlerts; } catch { }
                try { xlApp.EnableEvents = enableEvents; } catch { }
                try { xlApp.ScreenUpdating = screenUpdating; } catch { }
                if (wasSaved) try { wb.Saved = true; } catch { }
            }

            Process.Start(new ProcessStartInfo(pdfPath) { UseShellExecute = true });
            AppLogger.Info($"[BoardPrinter] Exported schedule PDF: {pdfPath}");
            SheetOperations.ShowStatusConfirmation($"Schedule PDF opened ({date:dd MMM yyyy}).");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[BoardPrinter] Print Schedule failed: {ex}");
            MessageBox.Show($"Could not export the schedule PDF.\n\n{ex.Message}", "Print Schedule", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
