using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>One production input as resolved under the root.</summary>
    public sealed class SwpProductionFile
    {
        public string Label = "";      // "Plant_Inputs.csv"
        public string Relative = "";   // the relative path found (or the first place looked)
        public string FullPath = "";   // "" when not found
        public bool Applicable = true; // false = this plant type never has it (never listed)
        public bool Found => FullPath.Length > 0;
    }

    /// <summary>Every production input the bridge consumes, resolved from one root folder.</summary>
    public sealed class SwpProductionFiles
    {
        public string Root = "";                   // the folder actually probed (after PRODUCTION descent)
        public SwpProductionFile Busfield = new SwpProductionFile { Label = "BusField.csv" };
        public SwpProductionFile PlantInputs = new SwpProductionFile { Label = "Plant_Inputs.csv" };
        public SwpProductionFile DestinationParams = new SwpProductionFile { Label = "Destination_Parameters.csv" };
        public SwpProductionFile BusExtendedDiagnostic = new SwpProductionFile { Label = "BUS_ExtendedDiagnostic.csv" };
        public SwpProductionFile Sensor = new SwpProductionFile { Label = "Sensor.csv" };
        public SwpProductionFile StrobeDevice = new SwpProductionFile { Label = "StrobeDevice.csv" };
        public List<string> AlarmStrings = new List<string>();
        public string AlarmStringsFolder = "";     // relative folder the Alarms_*.txt came from

        public IEnumerable<SwpProductionFile> Files
        {
            get
            {
                yield return Busfield;
                yield return PlantInputs;
                yield return DestinationParams;
                yield return BusExtendedDiagnostic;
                yield return Sensor;
                yield return StrobeDevice;
            }
        }

        /// <summary>The inputs this plant type consumes — the others are never mentioned.</summary>
        public IEnumerable<SwpProductionFile> ApplicableFiles
        {
            get
            {
                foreach (SwpProductionFile file in Files)
                    if (file.Applicable) yield return file;
            }
        }

        public string AlarmStringsPattern =>
            Path.Combine(AlarmStringsFolder, "Alarms_*.txt");

        /// <summary>The resolution as text, one line per applicable input (the build log).</summary>
        public List<string> ReportLines()
        {
            var lines = new List<string>();
            foreach (SwpProductionFile file in ApplicableFiles)
                lines.Add(file.Found
                    ? "✓  " + file.Relative
                    : "✗  " + file.Label + " not found — its checks stay off");
            lines.Add(AlarmStrings.Count > 0
                ? "✓  " + AlarmStringsPattern + " (" + AlarmStrings.Count.ToString(CultureInfo.InvariantCulture) + " files)"
                : "✗  Alarms_*.txt not found — the alarm-strings check stays off");
            return lines;
        }
    }

    /// <summary>
    /// Resolves the bridge's production inputs from ONE root folder by the layout the realtime
    /// applications actually use (measured on MSC, ISPC and upstream PRODUCTION trees):
    /// IODriver\ is the one constant; the rest is per plant — MSC\DB_Csv\ and
    /// DataDrawSupervisor\…\AlarmLists\ for the sorter, ISPC\db_csv\ and ISPC\AlarmLists\ for
    /// induction, ConveyorSys\SysGroupNN\Database\DB_Params|DB_Diag\ for upstream. Every input
    /// also accepts a flat copy directly under the root, and picking the backup folder that
    /// merely CONTAINS a PRODUCTION\ child descends into it. A missing file keeps its check off,
    /// exactly like leaving the old per-file picker empty did.
    /// </summary>
    public static class SwpProductionTree
    {
        public static SwpProductionFiles Resolve(string root, SwpPlantType plantType, int groupIndex)
        {
            var result = new SwpProductionFiles();
            root = (root ?? "").Trim();
            if (root.Length == 0 || !Directory.Exists(root)) return ApplyPlant(result, plantType);

            if (!Directory.Exists(Path.Combine(root, "IODriver")) &&
                Directory.Exists(Path.Combine(root, "PRODUCTION")))
                root = Path.Combine(root, "PRODUCTION");
            result.Root = root;

            ApplyPlant(result, plantType);
            Probe(root, result.Busfield, "IODriver\\BusField.csv", "BusField.csv");

            switch (plantType)
            {
                case SwpPlantType.Msc:
                    Probe(root, result.PlantInputs, "MSC\\DB_Csv\\Plant_Inputs.csv", "Plant_Inputs.csv");
                    Probe(root, result.DestinationParams, "MSC\\DB_Csv\\Destination_Parameters.csv", "Destination_Parameters.csv");
                    Probe(root, result.BusExtendedDiagnostic, "MSC\\DB_Csv\\BUS_ExtendedDiagnostic.csv", "BUS_ExtendedDiagnostic.csv");
                    ProbeAlarms(root, result, "DataDrawSupervisor\\DataDraw_PlantData\\AlarmLists", "");
                    break;

                case SwpPlantType.Ispc:
                    Probe(root, result.PlantInputs, "ISPC\\db_csv\\Plant_Inputs.csv", "Plant_Inputs.csv");
                    ProbeAlarms(root, result, "ISPC\\AlarmLists", "");
                    break;

                case SwpPlantType.Ups:
                {
                    string group = FindSysGroup(root, groupIndex) ?? "ConveyorSys\\SysGroup01";
                    Probe(root, result.PlantInputs, "Plant_Inputs.csv");   // upstream has none of its own
                    Probe(root, result.Sensor, group + "\\Database\\DB_Params\\Sensor.csv", "Sensor.csv");
                    Probe(root, result.StrobeDevice, group + "\\Database\\DB_Params\\StrobeDevice.csv", "StrobeDevice.csv");
                    ProbeAlarms(root, result, group + "\\Database\\DB_Diag", "");
                    break;
                }

                default:   // SAFETY: a Siemens PLC, no realtime tree — flat copies only
                    Probe(root, result.PlantInputs, "Plant_Inputs.csv");
                    ProbeAlarms(root, result, "");
                    break;
            }
            return result;
        }

        /// <summary>Which inputs the plant type consumes at all; the rest stay off the lists.</summary>
        private static SwpProductionFiles ApplyPlant(SwpProductionFiles result, SwpPlantType plantType)
        {
            bool sorter = plantType == SwpPlantType.Msc;
            bool upstream = plantType == SwpPlantType.Ups;
            result.DestinationParams.Applicable = sorter;
            result.BusExtendedDiagnostic.Applicable = sorter;
            result.Sensor.Applicable = upstream;
            result.StrobeDevice.Applicable = upstream;
            return result;
        }

        /// <summary>ConveyorSys\SysGroupNN — the configured group when present, else the only (or first) one.</summary>
        private static string FindSysGroup(string root, int groupIndex)
        {
            string conveyorSys = Path.Combine(root, "ConveyorSys");
            if (!Directory.Exists(conveyorSys)) return null;
            string wanted = "SysGroup" + groupIndex.ToString("00", CultureInfo.InvariantCulture);
            if (Directory.Exists(Path.Combine(conveyorSys, wanted))) return "ConveyorSys\\" + wanted;

            string[] groups = Directory.GetDirectories(conveyorSys, "SysGroup*");
            if (groups.Length == 0) return null;
            Array.Sort(groups, StringComparer.OrdinalIgnoreCase);
            return "ConveyorSys\\" + Path.GetFileName(groups[0]);
        }

        private static void Probe(string root, SwpProductionFile file, params string[] candidates)
        {
            foreach (string relative in candidates)
            {
                string full = Path.Combine(root, relative);
                if (File.Exists(full))
                {
                    file.Relative = relative;
                    file.FullPath = full;
                    return;
                }
            }
            file.Relative = candidates.Length > 0 ? candidates[0] : file.Label;
            file.FullPath = "";
        }

        private static void ProbeAlarms(string root, SwpProductionFiles result, params string[] folders)
        {
            foreach (string relative in folders)
            {
                string folder = relative.Length == 0 ? root : Path.Combine(root, relative);
                if (!Directory.Exists(folder)) continue;
                string[] files = Directory.GetFiles(folder, "Alarms_*.txt");
                if (files.Length == 0) continue;
                Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                result.AlarmStrings.AddRange(files);
                result.AlarmStringsFolder = relative;
                return;
            }
        }
    }
}
