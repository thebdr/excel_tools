using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// The migration proof. The declarative map must reproduce <see cref="SwpAddressMap"/> exactly
    /// for every shape the hardcoded mapper understood, before anything is allowed to depend on it.
    /// Generated mnemonics sweep the coordinate space; the real-file sweep is a separate test.
    ///
    /// Where the two DIFFER on purpose the difference is stated here as its own assertion, so an
    /// intended departure can never hide inside a parity failure.
    /// </summary>
    public class SwpMnemonicMapParityTests
    {
        private static readonly SwpMnemonicMap Map = SwpMnemonicMapDefaults.Official();

        private static SwpMapContext Context(IReadOnlyDictionary<string, int> chutes = null) =>
            new SwpMapContext
            {
                FirstLineId = 1,
                GroupIndex = 1,
                Rules = SwpIndexRules.Official(),
                Sources = chutes == null
                    ? null
                    : new Dictionary<string, IReadOnlyDictionary<string, int>>
                        { ["Destination_Parameters"] = chutes },
            };

        private static void AssertSame(string mnemonic, SwpPlantType plant,
            IReadOnlyDictionary<string, int> chutes = null)
        {
            SwpMappedMnemonic old = SwpAddressMap.Map(mnemonic, plant, firstLineId: 1,
                groupIndex: 1, deviceIndexOverride: null, chuteMap: chutes,
                rules: SwpIndexRules.Official());
            SwpMappedMnemonic now = Map.Map(mnemonic, plant, Context(chutes));

            Assert.True(old.Ok == now.Ok,
                mnemonic + ": old.Ok=" + old.Ok + " (" + old.Error + ") vs new.Ok=" + now.Ok +
                " (" + now.Error + ")");
            if (!old.Ok) return;

            Assert.Equal(old.Coordinate.SheetName, now.Coordinate.SheetName);
            Assert.Equal(old.Coordinate.TagType, now.Coordinate.TagType);
            Assert.Equal(old.Coordinate.TableEntry, now.Coordinate.TableEntry);
            Assert.Equal(old.Coordinate.Bit, now.Coordinate.Bit);
            Assert.Equal(old.IsWarning, now.IsWarning);
            Assert.Equal(old.ObjectType, now.ObjectType);
            Assert.Equal(old.RawDeviceNumber, now.RawDeviceNumber);
            Assert.Equal(old.ReferenceIndex, now.ReferenceIndex);
        }

        // ---- the format itself ---------------------------------------------------------------

        [Fact]
        public void Every_shipped_rule_survives_the_text_round_trip()
        {
            SwpMnemonicMap official = SwpMnemonicMapDefaults.Official();

            Assert.NotEmpty(official.Rules);
            Assert.All(official.Rules, r => Assert.Equal("", r.Error));
            // Re-serializing the parsed table must reproduce the source lines byte for byte.
            Assert.Equal(SwpMnemonicMapDefaults.OfficialLines(), official.ToLines());
        }

        [Fact]
        public void No_shipped_rule_uses_a_tab_or_a_pipe_separator_by_accident()
        {
            foreach (SwpMnemonicRule r in SwpMnemonicMapDefaults.Official().Rules)
                Assert.DoesNotContain("\t", r.Match);
        }

        // ---- MSC / SAFETY ----------------------------------------------------------------------

        [Theory]
        [InlineData(SwpPlantType.Msc)]
        [InlineData(SwpPlantType.Safety)]
        public void Cabinet_shapes_map_identically(SwpPlantType plant)
        {
            foreach (string aw in new[] { "AL", "WR" })
                foreach (int device in new[] { 0, 1, 19, 20, 255, 999 })
                    foreach (int bit in new[] { 0, 1, 31, 32, 33, 63, 64, 95 })
                        AssertSame("CABINET" + aw + device.ToString("000") + bit.ToString("00"), plant);
        }

        [Fact]
        public void Classdiag_shapes_map_identically_on_the_sorter()
        {
            foreach (int type in new[] { 62, 64, 66, 69, 73, 112, 113 })
                foreach (string aw in new[] { "AL", "WR" })
                    foreach (int device in new[] { 0, 7, 1234, 99999 })
                        foreach (int bit in new[] { 0, 31, 32, 63 })
                        {
                            string m = "CLASSDIAG" + aw + type.ToString("0000") +
                                device.ToString("00000") + bit.ToString("00");
                            var chutes = new Dictionary<string, int>
                            {
                                [device.ToString(CultureInfo.InvariantCulture)] = device + 5,
                            };
                            AssertSame(m, SwpPlantType.Msc, type == 62 ? chutes : null);
                        }
        }

        /// <summary>
        /// A safety SWP holds cabinets only (rules §9.2): the CLASSDIAG classes are sorter
        /// objects it has no sheet for, so under SAFETY both mappers refuse them and the table
        /// leaves them UNCLAIMED — a safety list's CLASSDIAG row is not a diagnostic at all rather
        /// than one bound for a sheet that does not exist.
        /// </summary>
        [Fact]
        public void Safety_refuses_classdiag_shapes_in_both()
        {
            foreach (int type in new[] { 62, 64, 66, 69, 73, 112, 113 })
                foreach (string aw in new[] { "AL", "WR" })
                {
                    string m = "CLASSDIAG" + aw + type.ToString("0000") + "00007" + "05";
                    var chutes = new Dictionary<string, int> { ["7"] = 12 };
                    AssertSame(m, SwpPlantType.Safety, type == 62 ? chutes : null);
                    Assert.False(SwpAddressMap.Map(m, SwpPlantType.Safety, chuteMap: chutes).Ok);
                    Assert.Null(Map.RuleFor(m, SwpPlantType.Safety));
                    Assert.NotNull(Map.RuleFor(m, SwpPlantType.Msc));   // the same shape IS a sorter diagnostic
                }
        }

        [Fact]
        public void Safety_rule_set_is_the_cabinet_family_and_nothing_else()
        {
            List<SwpMnemonicRule> safety = Map.Rules.Where(r => r.PlantType == SwpPlantType.Safety).ToList();
            List<SwpMnemonicRule> sorter = Map.Rules.Where(r => r.PlantType == SwpPlantType.Msc).ToList();

            Assert.NotEmpty(safety);
            Assert.All(safety, r => Assert.Equal(SwpRuleKind.Map, r.Kind));
            Assert.All(safety, r => Assert.True(r.ObjectKey == "CABINET" || r.MapsNowhere, r.Match));
            Assert.All(safety, r => Assert.True(r.Sheet == "CABINET" || r.MapsNowhere, r.Match));

            // Exactly the sorter's cabinet rows, same order and same regexes: the safety tooling
            // writes sorter syntax on purpose, so the two families must never drift apart.
            string[] safetyShapes = safety.OrderBy(r => r.Order).Select(r => r.Order + " " + r.Match).ToArray();
            string[] sorterCabinetShapes = sorter.Where(r => r.ObjectKey == "CABINET" || r.MapsNowhere)
                .OrderBy(r => r.Order).Select(r => r.Order + " " + r.Match).ToArray();
            Assert.Equal(sorterCabinetShapes, safetyShapes);
            Assert.Contains(safety, r => r.Match.Contains("EMERGENCY(?<i>"));
            Assert.DoesNotContain(safety, r => r.Match.Contains("CLASSDIAG"));
        }

        [Theory]
        [InlineData(SwpPlantType.Msc)]
        [InlineData(SwpPlantType.Safety)]
        public void Rubbish_fails_in_both(SwpPlantType plant)
        {
            foreach (string m in new[]
            {
                "", "   ", "CABINET", "CABINETXX00100", "CABINETAL0010",
                "CABINETAL001000", "CLASSDIAGAL006200001", "NOTAMNEMONIC",
                "CLASSDIAGAL999900001" + "00",
            })
                AssertSame(m, plant);
        }

        // ---- ISPC ------------------------------------------------------------------------------

        [Fact]
        public void Ispc_shapes_map_identically()
        {
            foreach (string aw in new[] { "AL", "WR" })
            {
                foreach (int line in new[] { 0, 1, 2, 9 })
                    foreach (int device in new[] { 0, 1, 123, 999 })
                        foreach (int bit in new[] { 0, 31, 32, 63 })
                            AssertSame("CABINET" + aw + line + device.ToString("000") +
                                bit.ToString("00"), SwpPlantType.Ispc);

                foreach (int line in new[] { 1, 2, 9 })
                    foreach (int device in new[] { 0, 1, 62, 99 })
                        foreach (int bit in new[] { 0, 31, 32 })
                            AssertSame("BELT" + aw + line + device.ToString("00") +
                                bit.ToString("00"), SwpPlantType.Ispc);

                foreach (int line in new[] { 1, 2, 9 })
                    foreach (int device in new[] { 0, 1, 9 })
                        foreach (int bit in new[] { 0, 31, 32 })
                            AssertSame("ZONE" + aw + line + device + bit.ToString("00"),
                                SwpPlantType.Ispc);
            }
        }

        // ---- UPS -------------------------------------------------------------------------------

        [Fact]
        public void Ups_shapes_map_identically()
        {
            foreach (string prefix in new[] { "ALLQ", "WARQ", "ALLB", "WARB", "ALLZ", "WARZ" })
                foreach (int system in new[] { 1, 2, 6 })
                    foreach (int device in new[] { 0, 1, 12, 99 })
                        foreach (int bit in new[] { 0, 31, 32, 63 })
                            AssertSame(
                                prefix + "_" + system + "_" + device.ToString("00") + "_" +
                                bit.ToString("00"),
                                SwpPlantType.Ups);
        }

        // ---- the deliberate departures ---------------------------------------------------------

        [Fact]
        public void Emergency_reaches_the_cabinet_without_being_renamed()
        {
            // The legacy path: rename to CABINETAL, then map. The new path maps EMERGENCY itself,
            // so the mnemonic that reaches Plant_Inputs is still the one the plant knows.
            SwpMappedMnemonic renamed = SwpAddressMap.Map("CABINETAL02000", SwpPlantType.Msc,
                rules: SwpIndexRules.Official());
            SwpMappedMnemonic direct = Map.Map("EMERGENCY02000", SwpPlantType.Msc, Context());

            Assert.True(direct.Ok, direct.Error);
            Assert.Equal(renamed.Coordinate.SheetName, direct.Coordinate.SheetName);
            Assert.Equal(renamed.Coordinate.TagType, direct.Coordinate.TagType);
            Assert.Equal(renamed.Coordinate.TableEntry, direct.Coordinate.TableEntry);
            Assert.Equal(renamed.Coordinate.Bit, direct.Coordinate.Bit);
            Assert.False(direct.IsWarning);
        }

        [Fact]
        public void Emergency_commands_are_recognized_and_go_nowhere()
        {
            foreach (string m in new[] { "EMERGENCY_ON", "EMERGENCY_RST" })
            {
                SwpMappedMnemonic mapped = Map.Map(m, SwpPlantType.Msc, Context());
                Assert.True(mapped.Ok);            // recognized, not an error
                Assert.Null(mapped.Coordinate);    // and deliberately without a destination
            }
        }

        [Fact]
        public void Only_the_zero_member_of_an_and_or_group_carries_it()
        {
            SwpMappedMnemonic first = Map.Map("ALLQ_AND_1_01_00_0", SwpPlantType.Ups, Context());
            Assert.True(first.Ok, first.Error);
            Assert.NotNull(first.Coordinate);
            Assert.Equal(0, first.Coordinate.Bit);

            // The legacy rewrite reached the same coordinate by mangling the name.
            SwpMappedMnemonic legacy = SwpAddressMap.Map("ALLQ_1_01_00", SwpPlantType.Ups,
                rules: SwpIndexRules.Official());
            Assert.Equal(legacy.Coordinate.SheetName, first.Coordinate.SheetName);
            Assert.Equal(legacy.Coordinate.TagType, first.Coordinate.TagType);
            Assert.Equal(legacy.Coordinate.TableEntry, first.Coordinate.TableEntry);

            foreach (string rest in new[] { "ALLQ_AND_1_01_00_1", "ALLQ_OR_1_01_00_7" })
            {
                SwpMappedMnemonic other = Map.Map(rest, SwpPlantType.Ups, Context());
                Assert.True(other.Ok);
                Assert.Null(other.Coordinate);
            }
        }

        [Fact]
        public void A_csv_backed_index_without_its_csv_is_an_error_not_a_guess()
        {
            // A DELIBERATE departure. The legacy mapper remapped chutes only "if chuteMap != null",
            // so a build that never loaded Destination_Parameters silently used the mnemonic's own
            // number — writing to whatever SWP row it happened to hit. A shape whose index comes
            // from a CSV cannot be placed without that CSV.
            SwpMappedMnemonic legacy = SwpAddressMap.Map("CLASSDIAGAL00620000107",
                SwpPlantType.Msc, chuteMap: null, rules: SwpIndexRules.Official());
            Assert.True(legacy.Ok);          // the old guess went through

            SwpMappedMnemonic now = Map.Map("CLASSDIAGAL00620000107", SwpPlantType.Msc, Context());
            Assert.False(now.Ok);
            Assert.Contains("Destination_Parameters", now.Error);

            // An EMPTY table is incoherent in the same way, not "loaded".
            SwpMappedMnemonic empty = Map.Map("CLASSDIAGAL00620000107", SwpPlantType.Msc,
                Context(new Dictionary<string, int>()));
            Assert.False(empty.Ok);

            // Loaded and coherent, but silent about THIS chute: still an error, and it says which.
            SwpMappedMnemonic missing = Map.Map("CLASSDIAGAL00620009907", SwpPlantType.Msc,
                Context(new Dictionary<string, int> { ["1"] = 4 }));
            Assert.False(missing.Ok);
            Assert.Contains("99", missing.Error);
        }

        [Fact]
        public void A_rule_that_does_not_parse_is_inert_and_never_fatal()
        {
            var log = new List<SwpLogEntry>();
            // Plant Order On Kind Object Match Emit Sheet Tag Word Index Bit Params Note
            SwpMnemonicMap map = SwpMnemonicMap.Parse(new[]
            {
                "\tMsc\t10\tx\tmap\tCABINET\t^CABINET(?<aw>AL|WR)(?<i>\\d{3}\t\tCABINET\tDB_X\tauto\t{i}\t{b}\t\t",
                "\tMsc\t20\tx\tmap\tCABINET\t^CABINET(?<aw>AL|WR)(?<i>\\d{3})(?<b>\\d{2})$\t\tCABINET\tDB_CABINETS_DIAG\tauto\t{i}\t{b}\t\t",
            }, log);

            Assert.Contains(log, l => l.Message.Contains("inert"));
            // The good rule below it still maps.
            Assert.True(map.Map("CABINETAL02000", SwpPlantType.Msc, Context()).Ok);
        }
    }
}
