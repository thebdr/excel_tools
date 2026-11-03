using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// One PRODUCTION root → every plant input, by the layouts measured on the real MSC, ISPC
    /// and upstream trees (IODriver\ constant; MSC\DB_Csv\ + DataDrawSupervisor AlarmLists;
    /// ISPC\db_csv\ + ISPC\AlarmLists; ConveyorSys\SysGroupNN\Database\DB_Params|DB_Diag).
    /// Inputs a plant type never has are not mentioned at all.
    /// </summary>
    public class SwpProductionTreeTests
    {
        private sealed class TempTree : IDisposable
        {
            public string Root { get; }

            public TempTree()
            {
                Root = Path.Combine(Path.GetTempPath(), "pt-swp-tree-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(Root);
            }

            public string Put(string relative)
            {
                string full = Path.Combine(Root, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(full)!);
                File.WriteAllText(full, "x");
                return full;
            }

            public void Dispose()
            {
                try { Directory.Delete(Root, true); } catch { }
            }
        }

        [Fact]
        public void Msc_tree_resolves_every_sorter_input_and_reports_it()
        {
            using (var tree = new TempTree())
            {
                string busfield = tree.Put(@"IODriver\BusField.csv");
                string plantInputs = tree.Put(@"MSC\DB_Csv\Plant_Inputs.csv");
                string destination = tree.Put(@"MSC\DB_Csv\Destination_Parameters.csv");
                string busExtended = tree.Put(@"MSC\DB_Csv\BUS_ExtendedDiagnostic.csv");
                tree.Put(@"MSC\DB_Csv\Cabinets.Csv");
                string alarmsIl = tree.Put(@"DataDrawSupervisor\DataDraw_PlantData\AlarmLists\Alarms_S1_IL01.txt");
                string alarmsCabinet = tree.Put(@"DataDrawSupervisor\DataDraw_PlantData\AlarmLists\Alarms_S1_CABINET.txt");
                tree.Put(@"DataDrawSupervisor\DataDraw_PlantData\AlarmLists\readme.txt");   // not an Alarms_ file

                SwpProductionFiles files = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Msc, 1);

                Assert.Equal(tree.Root, files.Root);
                Assert.Equal(busfield, files.Busfield.FullPath);
                Assert.Equal(@"IODriver\BusField.csv", files.Busfield.Relative);
                Assert.Equal(plantInputs, files.PlantInputs.FullPath);
                Assert.Equal(destination, files.DestinationParams.FullPath);
                Assert.Equal(busExtended, files.BusExtendedDiagnostic.FullPath);
                Assert.False(files.Sensor.Applicable);
                Assert.False(files.StrobeDevice.Applicable);
                Assert.Equal(new[] { "BusField.csv", "Plant_Inputs.csv", "Destination_Parameters.csv", "BUS_ExtendedDiagnostic.csv" },
                    files.ApplicableFiles.Select(f => f.Label).ToArray());
                Assert.Equal(new List<string> { alarmsCabinet, alarmsIl }, files.AlarmStrings);   // sorted
                Assert.Equal(@"DataDrawSupervisor\DataDraw_PlantData\AlarmLists", files.AlarmStringsFolder);

                Assert.Equal(new List<string>
                {
                    @"✓  IODriver\BusField.csv",
                    @"✓  MSC\DB_Csv\Plant_Inputs.csv",
                    @"✓  MSC\DB_Csv\Destination_Parameters.csv",
                    @"✓  MSC\DB_Csv\BUS_ExtendedDiagnostic.csv",
                    @"✓  DataDrawSupervisor\DataDraw_PlantData\AlarmLists\Alarms_*.txt (2 files)",
                }, files.ReportLines());
            }
        }

        [Fact]
        public void Ispc_tree_uses_the_induction_folders_and_never_mentions_sorter_or_upstream_inputs()
        {
            using (var tree = new TempTree())
            {
                string busfield = tree.Put(@"IODriver\BusField.csv");
                string plantInputs = tree.Put(@"ISPC\db_csv\Plant_Inputs.csv");
                string alarms = tree.Put(@"ISPC\AlarmLists\Alarms_S1_ILGRP_01.txt");
                tree.Put(@"MSC\DB_Csv\Destination_Parameters.csv");   // present but sorter-only

                SwpProductionFiles files = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Ispc, 1);

                Assert.Equal(busfield, files.Busfield.FullPath);
                Assert.Equal(plantInputs, files.PlantInputs.FullPath);
                Assert.Equal(@"ISPC\db_csv\Plant_Inputs.csv", files.PlantInputs.Relative);
                Assert.False(files.DestinationParams.Applicable);
                Assert.False(files.BusExtendedDiagnostic.Applicable);
                Assert.Equal("", files.DestinationParams.FullPath);
                Assert.Equal(new[] { "BusField.csv", "Plant_Inputs.csv" }, files.ApplicableFiles.Select(f => f.Label).ToArray());
                Assert.Equal(new List<string> { alarms }, files.AlarmStrings);
                Assert.Equal(@"ISPC\AlarmLists", files.AlarmStringsFolder);
                Assert.Equal(new List<string>
                {
                    @"✓  IODriver\BusField.csv",
                    @"✓  ISPC\db_csv\Plant_Inputs.csv",
                    @"✓  ISPC\AlarmLists\Alarms_*.txt (1 files)",
                }, files.ReportLines());
            }
        }

        [Fact]
        public void Ups_tree_prefers_the_configured_group_and_falls_back_to_the_first_one()
        {
            using (var tree = new TempTree())
            {
                tree.Put(@"IODriver\BusField.csv");
                tree.Put(@"IODriver_old\BusField.csv");   // ignored
                string sensor1 = tree.Put(@"ConveyorSys\SysGroup01\Database\DB_Params\Sensor.csv");
                string sensor2 = tree.Put(@"ConveyorSys\SysGroup02\Database\DB_Params\Sensor.csv");
                string strobe2 = tree.Put(@"ConveyorSys\SysGroup02\Database\DB_Params\StrobeDevice.csv");
                tree.Put(@"ConveyorSys\SysGroup02\Database\DB_Params\Belt.csv");
                string alarms2 = tree.Put(@"ConveyorSys\SysGroup02\Database\DB_Diag\Alarms_S1_UP04.txt");
                tree.Put(@"ConveyorSys\SysGroup02\Database\DB_Diag\BELT_DIAG.csv");
                string alarms1 = tree.Put(@"ConveyorSys\SysGroup01\Database\DB_Diag\Alarms_S1_UP01.txt");

                SwpProductionFiles group2 = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Ups, 2);
                Assert.Equal(sensor2, group2.Sensor.FullPath);
                Assert.Equal(@"ConveyorSys\SysGroup02\Database\DB_Params\Sensor.csv", group2.Sensor.Relative);
                Assert.Equal(strobe2, group2.StrobeDevice.FullPath);
                Assert.Equal(new List<string> { alarms2 }, group2.AlarmStrings);
                Assert.False(group2.DestinationParams.Applicable);
                Assert.False(group2.BusExtendedDiagnostic.Applicable);
                Assert.True(group2.PlantInputs.Applicable);
                Assert.False(group2.PlantInputs.Found);   // upstream has no Plant_Inputs — flat copy only
                Assert.Contains("✗  Plant_Inputs.csv not found — its checks stay off", group2.ReportLines());
                Assert.DoesNotContain(group2.ReportLines(), line => line.Contains("Destination_Parameters"));

                SwpProductionFiles group9 = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Ups, 9);
                Assert.Equal(sensor1, group9.Sensor.FullPath);           // no SysGroup09 → first group
                Assert.False(group9.StrobeDevice.Found);                 // group 1 has none
                Assert.Equal(new List<string> { alarms1 }, group9.AlarmStrings);
            }
        }

        [Fact]
        public void A_backup_folder_descends_into_its_production_child()
        {
            using (var tree = new TempTree())
            {
                string busfield = tree.Put(@"PRODUCTION\IODriver\BusField.csv");
                string plantInputs = tree.Put(@"PRODUCTION\MSC\DB_Csv\Plant_Inputs.csv");
                tree.Put(@"somefile.txt");

                SwpProductionFiles files = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Msc, 1);

                Assert.Equal(Path.Combine(tree.Root, "PRODUCTION"), files.Root);
                Assert.Equal(busfield, files.Busfield.FullPath);
                Assert.Equal(plantInputs, files.PlantInputs.FullPath);
            }
        }

        [Fact]
        public void Flat_copies_directly_under_the_root_are_accepted()
        {
            using (var tree = new TempTree())
            {
                string busfield = tree.Put("BusField.csv");
                string plantInputs = tree.Put("Plant_Inputs.csv");
                string alarms = tree.Put("Alarms_S1_CABINET.txt");

                SwpProductionFiles files = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Msc, 1);

                Assert.Equal(busfield, files.Busfield.FullPath);
                Assert.Equal("BusField.csv", files.Busfield.Relative);
                Assert.Equal(plantInputs, files.PlantInputs.FullPath);
                Assert.False(files.DestinationParams.Found);
                Assert.Equal(new List<string> { alarms }, files.AlarmStrings);
                Assert.Equal("", files.AlarmStringsFolder);
                Assert.Contains("✓  Alarms_*.txt (1 files)", files.ReportLines());
            }
        }

        /// <summary>
        /// A Siemens safety PLC has no IODriver busfield and no realtime input DB: the SCADA's
        /// alarm export is its ONE production input. The FVT safety system keeps it in an
        /// "Alarms List" folder beside the SWP, so the SWP's version folder is the root.
        /// </summary>
        [Fact]
        public void Safety_consumes_the_alarm_export_only()
        {
            using (var tree = new TempTree())
            {
                tree.Put(@"IODriver\BusField.csv");          // the sorter's — never the safety PLC's
                tree.Put("Plant_Inputs.csv");                // likewise
                tree.Put("SWP_01_Safety.xlsx");
                string alarms = tree.Put(@"Alarms List\Alarms_S1_CABINET.txt");

                SwpProductionFiles files = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Safety, 1);

                Assert.Empty(files.ApplicableFiles);
                Assert.False(files.Busfield.Applicable);
                Assert.False(files.PlantInputs.Applicable);
                Assert.Equal("", files.Busfield.FullPath);
                Assert.Equal(new List<string> { alarms }, files.AlarmStrings);
                Assert.Equal("Alarms List", files.AlarmStringsFolder);
                Assert.Equal(new[] { @"✓  Alarms List\Alarms_*.txt (1 files)" }, files.ReportLines());
            }
        }

        [Fact]
        public void Safety_accepts_the_sorter_style_AlarmLists_folder_and_a_flat_copy()
        {
            using (var tree = new TempTree())
            {
                string alarms = tree.Put(@"AlarmLists\Alarms_S1_CABINET.txt");
                SwpProductionFiles files = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Safety, 1);
                Assert.Equal(new List<string> { alarms }, files.AlarmStrings);
                Assert.Equal("AlarmLists", files.AlarmStringsFolder);
            }

            using (var tree = new TempTree())
            {
                string alarms = tree.Put("Alarms_S1_CABINET.txt");
                SwpProductionFiles files = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Safety, 1);
                Assert.Equal(new List<string> { alarms }, files.AlarmStrings);
                Assert.Equal("", files.AlarmStringsFolder);
            }

            using (var tree = new TempTree())
            {
                // The sorter's own export folder is NOT the safety PLC's.
                tree.Put(@"DataDrawSupervisor\DataDraw_PlantData\AlarmLists\Alarms_S1_CABINET.txt");
                SwpProductionFiles files = SwpProductionTree.Resolve(tree.Root, SwpPlantType.Safety, 1);
                Assert.Empty(files.AlarmStrings);
                Assert.Equal(new[] { "✗  Alarms_*.txt not found — the alarm-strings check stays off" }, files.ReportLines());
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        [InlineData(@"C:\definitely\not\here\pt-swp-tree-missing")]
        public void A_missing_root_resolves_nothing_and_says_so(string root)
        {
            SwpProductionFiles files = SwpProductionTree.Resolve(root, SwpPlantType.Msc, 1);

            Assert.Equal("", files.Root);
            foreach (SwpProductionFile file in files.Files) Assert.False(file.Found);
            Assert.Empty(files.AlarmStrings);
            List<string> lines = files.ReportLines();
            Assert.Equal(5, lines.Count);   // the four sorter inputs + the alarm strings
            Assert.Equal("✗  BusField.csv not found — its checks stay off", lines[0]);
            Assert.Equal("✗  Alarms_*.txt not found — the alarm-strings check stays off", lines[lines.Count - 1]);
        }

        [Fact]
        public void The_sidecar_keeps_the_production_root_and_the_resolved_bus_extended_path()
        {
            var document = new SwpSidecarDocument();
            document.Settings.ProductionRoot = @"C:\backup\FVZ_MSC_20260807_1617\PRODUCTION";
            document.Settings.BusExtendedDiagnosticCsvPath = @"C:\backup\FVZ_MSC_20260807_1617\PRODUCTION\MSC\DB_Csv\BUS_ExtendedDiagnostic.csv";
            string json = SwpSidecarJson.Serialize(document);

            const string property = "\"productionRoot\": \"C:\\\\backup\\\\FVZ_MSC_20260807_1617\\\\PRODUCTION\"";
            Assert.Contains(property, json);
            Assert.Contains("\"busExtendedDiagnosticCsvPath\": ", json);
            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument back, out string error), error);
            Assert.Equal(@"C:\backup\FVZ_MSC_20260807_1617\PRODUCTION", back.Settings.ProductionRoot);
            Assert.Equal(document.Settings.BusExtendedDiagnosticCsvPath, back.Settings.BusExtendedDiagnosticCsvPath);

            // A sidecar from before the root reads back with the root empty.
            Assert.True(SwpSidecarJson.TryParse(json.Replace(property + ",", ""), out SwpSidecarDocument legacy, out error), error);
            Assert.Equal("", legacy.Settings.ProductionRoot);
        }
    }
}
