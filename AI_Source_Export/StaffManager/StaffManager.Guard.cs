using ExcelDna.Integration;

public static class Guard
{
    public static bool IsStaffWorkbook(dynamic? wb = null)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic targetWb = wb ?? xlApp.ActiveWorkbook;
            if (targetWb == null) return false;

            // Check for signature sheet and signature ListObject table
            dynamic ws = targetWb.Sheets["All Staffs"];
            if (ws != null && ws.ListObjects["AllStaffsTable"] != null)
            {
                return true;
            }
        }
        catch { }
        return false;
    }
}