using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpAddressMapTests
    {
        // ---- MSC (sorter) ----

        [Theory]
        [InlineData("CABINETAL00100", "CABINET|DB_CABINETS_DIAG.Alarm1>1>0", false)]
        [InlineData("CABINETAL00131", "CABINET|DB_CABINETS_DIAG.Alarm1>1>31", false)]
        [InlineData("CABINETAL00132", "CABINET|DB_CABINETS_DIAG.Alarm2>1>32", false)]
        [InlineData("CABINETWR00553", "CABINET|DB_CABINETS_DIAG.Warning2>5>53", true)]
        [InlineData("CABINETAL01200", "CABINET|DB_CABINETS_DIAG.Alarm1>12>0", false)]
        public void Msc_cabinet_maps(string mnemonic, string expected, bool warning)
        {
            SwpMappedMnemonic r = SwpAddressMap.Map(mnemonic, SwpPlantType.Msc);
            Assert.True(r.Ok, r.Error);
            Assert.Equal(expected, r.Coordinate.FullKey);
            Assert.Equal(warning, r.IsWarning);
            Assert.Equal("CABINET", r.ObjectType);
        }

        [Theory]
        [InlineData("0062", "CHUTE", "DB_CHUTE")]
        [InlineData("0064", "SCANNER", "DB_SCANNER")]
        [InlineData("0066", "CONTROL_ZONE", "DB_CONTROLZONE")]
        [InlineData("0069", "CHECK_POSITION", "DB_CHECKPOSITION")]
        [InlineData("0073", "CELL_TEST", "DB_CELLTEST")]
        public void Msc_classdiag_types_map(string typeDigits, string sheet, string tag)
        {
            string mnemonic = "CLASSDIAGAL" + typeDigits + "00003" + "07";
            SwpMappedMnemonic r = SwpAddressMap.Map(mnemonic, SwpPlantType.Msc);
            Assert.True(r.Ok, r.Error);
            Assert.Equal(sheet + "|" + tag + ".Alarm1>3>7", r.Coordinate.FullKey);
        }

        [Fact]
        public void Msc_unknown_classdiag_type_fails()
        {
            SwpMappedMnemonic r = SwpAddressMap.Map("CLASSDIAGAL00990000107", SwpPlantType.Msc);
            Assert.False(r.Ok);
            Assert.Contains("MSC-1", r.Error);
        }

        [Fact]
        public void Msc_chute_resolves_through_destination_parameters_order()
        {
            var chuteMap = new Dictionary<string, int> { { "3", 17 } };
            SwpMappedMnemonic r = SwpAddressMap.Map("CLASSDIAGAL00620000307", SwpPlantType.Msc, chuteMap: chuteMap);
            Assert.True(r.Ok, r.Error);
            Assert.Equal(17, r.Coordinate.TableEntry);
            Assert.Equal(3, r.RawDeviceNumber);
        }

        [Fact]
        public void Msc_chute_missing_from_destination_parameters_fails_explicitly()
        {
            var chuteMap = new Dictionary<string, int> { { "9", 0 } };
            SwpMappedMnemonic r = SwpAddressMap.Map("CLASSDIAGAL00620000307", SwpPlantType.Msc, chuteMap: chuteMap);
            Assert.False(r.Ok);
            Assert.Contains("Destination_Parameters", r.Error);
        }

        [Fact]
        public void Msc_non_chute_ignores_chute_map()
        {
            var chuteMap = new Dictionary<string, int> { { "3", 17 } };
            SwpMappedMnemonic r = SwpAddressMap.Map("CLASSDIAGAL00640000307", SwpPlantType.Msc, chuteMap: chuteMap);
            Assert.True(r.Ok, r.Error);
            Assert.Equal(3, r.Coordinate.TableEntry);
        }

        [Theory]
        [InlineData("CABINETXX00100", "MSC-3")]   // neither AL nor WR
        [InlineData("CABINETAL0010", "MSC-2")]    // wrong length
        [InlineData("SOMETHINGELSE1", "MSC-2")]   // wrong shape entirely
        [InlineData("CLASSDIAGXX00620000107", "MSC-3")] // shape fits, but neither AL nor WR
        [InlineData("CLASSDIAGAL0062000010X", "MSC-2")] // digit block broken
        public void Msc_invalid_shapes_fail(string mnemonic, string code)
        {
            SwpMappedMnemonic r = SwpAddressMap.Map(mnemonic, SwpPlantType.Msc);
            Assert.False(r.Ok);
            Assert.Contains(code, r.Error);
        }

        [Fact]
        public void Device_index_override_wins_everywhere()
        {
            Assert.Equal(7, SwpAddressMap.Map("CABINETAL00100", SwpPlantType.Msc, deviceIndexOverride: 7).Coordinate.TableEntry);
            Assert.Equal(7, SwpAddressMap.Map("CABINETAL000162", SwpPlantType.Ispc, deviceIndexOverride: 7).Coordinate.TableEntry);
            Assert.Equal(7, SwpAddressMap.Map("ALLB_1_001_50", SwpPlantType.Ups, deviceIndexOverride: 7).Coordinate.TableEntry);
        }

        // ---- ISPC (induction) ----

        [Fact]
        public void Ispc_group_cabinet_line_zero_goes_to_the_group_sheet()
        {
            SwpMappedMnemonic r = SwpAddressMap.Map("CABINETAL000162", SwpPlantType.Ispc, firstLineId: 1, groupIndex: 2);
            Assert.True(r.Ok, r.Error);
            Assert.Equal("ILGRP02_CABINET|DB_ILGRP02_CABINET.Alarm2>0>62", r.Coordinate.FullKey);
            Assert.Equal(0, r.ReferenceIndex);
            Assert.Equal(1, r.RawDeviceNumber);
        }

        [Fact]
        public void Ispc_line_cabinet_offsets_by_first_line_id()
        {
            SwpMappedMnemonic r = SwpAddressMap.Map("CABINETAL200262", SwpPlantType.Ispc, firstLineId: 3, groupIndex: 1);
            Assert.True(r.Ok, r.Error);
            Assert.Equal("IL04_CABINET|DB_IL04_CABINET.Alarm2>1>62", r.Coordinate.FullKey);
            Assert.Equal(2, r.ReferenceIndex);
        }

        [Theory]
        [InlineData("BELTAL10162", "IL01_BELT|DB_IL01_BELT.Alarm2>0>62", "BELT")]
        [InlineData("ZONEWR1204", "IL01_ZONE|DB_IL01_ZONE.Warning1>1>4", "ZONE")]
        public void Ispc_belt_and_zone_map(string mnemonic, string expected, string objectType)
        {
            SwpMappedMnemonic r = SwpAddressMap.Map(mnemonic, SwpPlantType.Ispc, firstLineId: 1, groupIndex: 1);
            Assert.True(r.Ok, r.Error);
            Assert.Equal(expected, r.Coordinate.FullKey);
            Assert.Equal(objectType, r.ObjectType);
        }

        [Fact]
        public void Ispc_foreign_shape_fails()
        {
            Assert.False(SwpAddressMap.Map("ALLB_1_001_50", SwpPlantType.Ispc).Ok);
        }

        // ---- UPS (upstream) ----

        [Fact]
        public void Ups_belt_alarm_matches_the_real_list_sample()
        {
            // Golden sample read from the legacy tool's own List sheet.
            SwpMappedMnemonic r = SwpAddressMap.Map("ALLB_1_001_50", SwpPlantType.Ups, firstLineId: 1);
            Assert.True(r.Ok, r.Error);
            Assert.Equal("UP01_BELT|DB_UP01_BELT.Alarm2>0>50", r.Coordinate.FullKey);
            Assert.True(r.IsBeltAlarm);
            Assert.Equal(1, r.ReferenceIndex);
            Assert.Equal(1, r.RawDeviceNumber);
        }

        [Theory]
        [InlineData("WARQ_2_01_00", 5, "UP06_CABINET|DB_UP06_CABINET.Warning1>0>0", true)]
        [InlineData("ALLZ_1_02_12", 1, "UP01_ZONE|DB_UP01_ZONE.Alarm1>1>12", false)]
        [InlineData("ALLQ_1_01_07", 1, "UP01_CABINET|DB_UP01_CABINET.Alarm1>0>7", false)]
        public void Ups_variants_map(string mnemonic, int firstLine, string expected, bool warning)
        {
            SwpMappedMnemonic r = SwpAddressMap.Map(mnemonic, SwpPlantType.Ups, firstLineId: firstLine);
            Assert.True(r.Ok, r.Error);
            Assert.Equal(expected, r.Coordinate.FullKey);
            Assert.Equal(warning, r.IsWarning);
        }

        [Theory]
        [InlineData("ALLX_1_1_0", "UP-2")]  // unknown object suffix
        [InlineData("XLLB_1_1_0", "UP-1")]  // neither alarm nor warning
        [InlineData("ALLB_1_001", "UP-1")]  // too few parts
        [InlineData("ALLB_1_X_0", "UP-1")]  // non-numeric device
        public void Ups_invalid_shapes_fail(string mnemonic, string code)
        {
            SwpMappedMnemonic r = SwpAddressMap.Map(mnemonic, SwpPlantType.Ups);
            Assert.False(r.Ok);
            Assert.Contains(code, r.Error);
        }

        [Fact]
        public void Empty_mnemonic_fails()
        {
            Assert.False(SwpAddressMap.Map("  ", SwpPlantType.Msc).Ok);
        }
    }
}
