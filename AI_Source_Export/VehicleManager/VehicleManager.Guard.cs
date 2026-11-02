namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;

public static class Guard
{
    public static bool IsVehicleWorkbook(dynamic? wb = null)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic targetWb = wb ?? xlApp.ActiveWorkbook;
            if (targetWb == null) return false;

            dynamic ws = null;
            try { ws = targetWb.Sheets["All Vehicles"]; } catch { }
            if (ws == null)
            {
                try { ws = targetWb.Sheets["AllVehicles"]; } catch { }
            }
            if (ws == null) return false;

            dynamic table = null;
            try { table = ws.ListObjects["AllVehiclesTable"]; } catch { }
            return table != null;
        }
        catch
        {
            return false;
        }
    }
}