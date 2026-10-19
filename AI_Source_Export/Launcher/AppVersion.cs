using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKSB_App.Launcher
{
    public sealed class AppVersion
    {
        public string Version { get; set; } = "1.0.0";
        public DateTime LastUpdatedUtc { get; set; } = DateTime.UtcNow;
        public string Environment { get; set; } = "Production";

        /// <summary>
        /// Optional hash map for file integrity and granular incremental diffs.
        /// Key: Relative file path (e.g., "AddIns/Portal-AddIn64.xll")
        /// Value: SHA256 checksum
        /// </summary>
        public Dictionary<string, string> FileHashes { get; set; } = new();
    }
}
