using System.Collections.Generic;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// Seed data for a system with no sidecar yet. The dictionaries start empty (the legacy
    /// tool's Description Dictionary sheet held only placeholder rows); the strobe templates are
    /// the verbatim harvest of its StrobeDevices sheet — including its quirks (the "LENGHT"
    /// spelling and the duplicated _01 bit stand exactly as the plant knows them). Everything is
    /// editable on the workspace sheets afterwards.
    /// </summary>
    public static class SwpDefaults
    {
        public static List<SwpDictionaryEntry> EmptyDictionary() => new List<SwpDictionaryEntry>();

        public static List<SwpStrobeTemplate> DefaultStrobeTemplates()
        {
            var scale = new SwpStrobeTemplate { TypeId = "1", TypeDescription = "Scale" };
            AddLines(scale,
                "WARQ_%sys%_%cabinet%_01|SCALE CERTIFICATION PROCEDURE ONGOING",
                "WARQ_%sys%_%cabinet%_06|Fault Received from Scale",
                "WARQ_%sys%_%cabinet%_09|Program run from scale not received",
                "WARQ_%sys%_%cabinet%_10|Scale cabinet switch tripped",
                "WARQ_%sys%_%cabinet%_20|Scale bypass activated",
                "WARQ_%sys%_%cabinet%_21|Scale auto tare failed",
                "WARQ_%sys%_%cabinet%_38|Scale not ready",
                "ALLQ_%sys%_%cabinet%_01|SCALE CERTIFICATION PROCEDURE ONGOING",
                "ALLQ_%sys%_%cabinet%_06|Fault Received from Scale",
                "ALLQ_%sys%_%cabinet%_09|Program run from scale not received",
                "ALLQ_%sys%_%cabinet%_10|Scale cabinet switch tripped",
                "ALLQ_%sys%_%cabinet%_20|Scale bypass activated",
                "ALLQ_%sys%_%cabinet%_21|Scale auto tare failed",
                "ALLQ_%sys%_%cabinet%_38|Scale not ready");

            var bottomCamera = new SwpStrobeTemplate { TypeId = "2", TypeDescription = "BottomCamera" };
            string[] cameraBits =
            {
                "00|Too many answer timeouts",
                "02|DATALOGIC BOX TOL FAIL LENGHT",
                "03|DATALOGIC BOX TOL FAIL WIDTH",
                "04|DATALOGIC BOX TOL FAIL HEIGHT",
                "05|DATALOGIC BOX BELT STABILITY",
                "06|DATALOGIC NODE RESET",
                "07|DATALOGIC DIGITIZER MOTOR FAILURE",
                "08|DATALOGIC DIGITIZER LASER FAILURE",
                "09|DATALOGIC BOX SIDE BY SIDE",
                "10|DATALOGIC BOX NO DATA",
                "11|DATALOGIC INPUT 1 FAILURE",
                "12|DataLogic Camera Encoder Failure",
                "13|DATALOGIC SD CARD FAILURE",
                "14|DATALOGIC PROTOCOL INDEX ALARM ON MAIN PORT",
                "15|DATALOGIC PROTOCOL INDEX ALARM ON AUX PORT",
                "16|PROTOCOL INDEX ALARM ON USER SOCKET 1 PORT",
                "17|PROTOCOL INDEX ALARM ON USER SOCKET 2 PORT",
                "18|PROTOCOL INDEX ALARM ON USER SOCKET 3 PORT",
                "19|PROTOCOL INDEX ALARM ON USER SOCKET 4 PORT",
                "20|PROTOCOL INDEX ALARM ON USER SOCKET 5 PORT",
                "21|PROTOCOL INDEX ALARM ON USER SOCKET 6 PORT",
                "22|DATALOGIC COMMUNICATOR ERROR",
                "23|DATALOGIC NET RING OPEN",
                "24|DATALOGIC HYBRID CAMERA SLAVE INDICATION",
                "25|DATALOGIC ENGINE OVERLOAD",
                "26|DATALOGIC TOO NEAR",
                "27|DATALOGIC TOO NEAR ERROR",
                "28|DATALOGIC CAMERA VIDEO COMMUNIC",
                "29|DATALOGIC SERIAL COMMUNIC DEVICE",
                "30|DATALOGIC CAMERA NOT RUNNING",
                "01|DATALOGIC NO IMAGE CAPTURED WHEN TRIGGERED",
                "32|DATALOGIC ILLUM OVERTEMP",
                "33|DATALOGIC CAMERA AUTOFOCUS",
                "34|DATALOGIC HEIGHT SENSOR",
                "35|DATALOGIC HOST SOCKET PORT",
                "36|DATALOGIC UNABLE OPEN SERIAL AUX PORT",
                "37|DATALOGIC UNABLE OPEN USER SOCKET",
                "38|DATALOGIC UNABLE OPEN SERIAL SCANNER CLUSTER",
                "39|UNABLE OPEN SERIAL AUX SCANNER CLUSTER",
                "40|DATALOGIC UNABLE OPEN SOCKET SCANNER CLUSTER",
                "41|DATALOGIC CAMERA PARAM NOT LOADED",
                "42|DATALOGIC RIGHT SENSOR",
                "43|DATALOGIC LEFT SENSOR",
                "44|DATALOGIC ILLUM COMMUNICATOR FAILURE",
                "45|DATALOGIC ILLUM LED BAR FAILURE",
                "46|DATALOGIC ILLUM TEMPERATURE FAILURE",
                "47|DATALOGIC ILLUM FAN1 FAILURE",
                "48|DATALOGIC ILLUM FAN2 FAILURE",
                "49|DATALOGIC CCD COMMUNICATION FAILURE",
                "50|DATALOGIC CCD OVERTEMP",
                "51|DATALOGIC CAMERA CONTROLLER OVERTEMP",
                "52|DATALOGIC CAMERA INTERNAL FAN ERROR",
                "53|DATALOGIC CAMERA EXTERNAL FAN1 ERROR",
                "54|DATALOGIC CAMERA EXTERNAL FAN2 ERROR",
                "55|DATALOGIC BACKUP COMPACT FLSH MISSING",
                "56|DATALOGIC INCOMPATIBLE SPACE CONFIG",
                "57|DATALOGIC PROTOCOL INDEX MASTER",
                "58|DATALOGIC ENCODER FAILURE MASTER",
                "59|DATALOGIC COMMUNICATION NVS SC5000",
                "60|DATALOGIC GENERIC FAILURE",
            };
            foreach (string prefix in new[] { "ALLQ", "WARQ" })
                foreach (string line in cameraBits)
                {
                    int split = line.IndexOf('|');
                    bottomCamera.Mnemonics.Add(prefix + "_%sys%_%cabinet%_" + line.Substring(0, split));
                    bottomCamera.Descriptions.Add(line.Substring(split + 1));
                }

            return new List<SwpStrobeTemplate> { scale, bottomCamera };
        }

        private static void AddLines(SwpStrobeTemplate template, params string[] lines)
        {
            foreach (string line in lines)
            {
                int split = line.IndexOf('|');
                template.Mnemonics.Add(line.Substring(0, split));
                template.Descriptions.Add(line.Substring(split + 1));
            }
        }
    }
}
