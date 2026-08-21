using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// Card sheets are all in by default; only the deliberately unticked ones are remembered,
    /// so a card added to the IO list later starts ticked. Additive property — older sidecars
    /// read back with nothing excluded.
    /// </summary>
    public class SwpSidecarCardSheetsTests
    {
        [Fact]
        public void Excluded_card_sheets_round_trip_beside_the_ticked_ones()
        {
            var document = new SwpSidecarDocument();
            document.Settings.CardSheets.Add("R1.S1.C1");
            document.Settings.CardSheets.Add("R1.S1.C3");
            document.Settings.ExcludedCardSheets.Add("R1.S1.C2");
            string json = SwpSidecarJson.Serialize(document);

            Assert.Contains("\"excludedCardSheets\": [", json);
            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument back, out string error), error);
            Assert.Equal(new[] { "R1.S1.C1", "R1.S1.C3" }, back.Settings.CardSheets);
            Assert.Equal(new[] { "R1.S1.C2" }, back.Settings.ExcludedCardSheets);
        }

        [Fact]
        public void A_sidecar_without_the_property_excludes_nothing()
        {
            var document = new SwpSidecarDocument();
            document.Settings.CardSheets.Add("R1.S1.C1");
            string json = SwpSidecarJson.Serialize(document);
            int start = json.IndexOf("\"excludedCardSheets\"");
            int end = json.IndexOf(']', start) + 1;
            string legacy = json.Remove(start, end - start).Replace(",\r\n\r\n", "\r\n");
            legacy = System.Text.RegularExpressions.Regex.Replace(legacy, @",(\s*),", ",$1");

            Assert.True(SwpSidecarJson.TryParse(legacy, out SwpSidecarDocument back, out string error), error);
            Assert.Empty(back.Settings.ExcludedCardSheets);
            Assert.Equal(new[] { "R1.S1.C1" }, back.Settings.CardSheets);
        }
    }
}
