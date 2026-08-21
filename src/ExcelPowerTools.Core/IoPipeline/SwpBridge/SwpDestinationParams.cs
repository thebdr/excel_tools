using System.Collections.Generic;
using System.Globalization;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// Destination_Parameters.csv (next to the plant CSVs) lists the sorter chutes in SWP order.
    /// A chute mnemonic carries the MSC index, but the SWP CHUTE sheet is laid out in file order —
    /// so the device index is the chute's zero-based ordinal among the active ("@") lines, exactly
    /// as the legacy GetChuteSwpDeviceIndex counted it. Field 1 of each "@" line is the MSC index.
    /// </summary>
    public static class SwpDestinationParams
    {
        public static Dictionary<string, int> ParseChuteMap(IEnumerable<string> lines, List<SwpLogEntry> log)
        {
            var map = new Dictionary<string, int>();
            if (lines == null) return map;

            int ordinal = 0;
            int lineNumber = 0;
            foreach (string line in lines)
            {
                lineNumber++;
                if (line == null || line.Length == 0 || line[0] != '@') continue;

                char separator = line.Length > 1 ? line[1] : '\0';
                if (separator != ',' && separator != ';')
                {
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                        "Destination_Parameters line " + lineNumber + ": separator '" + separator + "' is invalid"));
                    // The legacy tool logged and split with it anyway; an unsplittable line simply
                    // yields no index and the ordinal still advances (it is an active chute line).
                }

                string[] fields = line.Split(separator);
                if (fields.Length > 1)
                {
                    string key = IoListLogic.Norm(fields[1]);
                    if (key.Length > 0 && !map.ContainsKey(key)) map[key] = ordinal;
                }
                ordinal++;
            }
            return map;
        }

        /// <summary>The chuteMap key for an MSC chute index parsed out of a mnemonic.</summary>
        public static string Key(int mscIndex) => mscIndex.ToString(CultureInfo.InvariantCulture);
    }
}
