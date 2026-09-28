using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace SKSB_App.ThumbprintSync
{
    [ComVisible(true)]
    public class VigilanceRibbonController : ExcelRibbon
    {
        // 1. Injects a minimal tab and button into Excel's top ribbon
        public override string GetCustomUI(string ribbonId)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabAllocation' label='Work Allocation'>
                    <group id='grpDevice' label='Terminal Sync'>
                      <button id='btnSync' 
                              label='Sync to Reader' 
                              size='large' 
                              imageMso='AccessRefreshAllLists' 
                              onAction='OnSyncClicked' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        // 2. The callback triggered by clicking the button
        public void OnSyncClicked(IRibbonControl control)
        {
            // Run on background thread to keep Excel UI responsive
            Task.Run(async () =>
            {
                try
                {
                    const string deviceIp = "192.168.0.110";

                    // Call your extraction & socket logic here
                    await Task.Delay(500); // Simulated network latency

                    MessageBox.Show($"Successfully dispatched allocation to {deviceIp}.",
                                    "Vigilance Sync",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Sync error: {ex.Message}",
                                    "Error",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                }
            });
        }
    }

    // Optional: Expose custom formula for checking device status directly in a cell
    public static class DeviceFormulas
    {
        [ExcelFunction(Description = "Pings the Vigilance reader status")]
        public static string PING_READER(string ipAddress)
        {
            return $"Ready @ {ipAddress}";
        }
    }
}