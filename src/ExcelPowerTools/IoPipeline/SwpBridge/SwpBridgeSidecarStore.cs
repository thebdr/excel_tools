using System;
using System.IO;
using System.Text;
using ExcelPowerTools.Core.Files;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// The sidecar's disk home: swp-bridge.json in the system's SWP ROOT folder — one level above
    /// the dated version folders, so it survives every version hop and there is exactly one per
    /// system. Written atomically (with .bak) on every decision.
    /// </summary>
    internal static class SwpBridgeSidecarStore
    {
        public const string FileName = "swp-bridge.json";

        public static string PathFor(string swpRootFolder) => Path.Combine(swpRootFolder, FileName);

        /// <summary>
        /// Loads the sidecar of an SWP root. Null when none exists; throws with a recovery hint
        /// when one exists but cannot be read — silently dropping decisions is never acceptable.
        /// </summary>
        public static SwpSidecarDocument Load(string swpRootFolder)
        {
            string path = PathFor(swpRootFolder);
            SourceReadState state = AtomicFileStore.ReadSource(path, out byte[] bytes, out string readError);
            if (state == SourceReadState.Missing) return null;
            if (state == SourceReadState.Unavailable)
                throw new InvalidOperationException("The SWP Bridge sidecar could not be read:\n" + path + "\n" + readError);

            string json = Encoding.UTF8.GetString(bytes);
            if (json.Length > 0 && json[0] == (char)0xFEFF) json = json.Substring(1);

            if (!SwpSidecarJson.TryParse(json, out SwpSidecarDocument document, out string parseError))
                throw new InvalidOperationException(
                    "The SWP Bridge sidecar is damaged and was not loaded:\n" + path + "\n" + parseError +
                    "\n\nA previous version may exist as \"" + FileName + ".bak\".");
            return document;
        }

        public static void Save(string swpRootFolder, SwpSidecarDocument document)
        {
            document.GeneratedUtc = DateTime.UtcNow.ToString("O");
            string json = SwpSidecarJson.Serialize(document);
            AtomicFileStore.WriteWithBackup(PathFor(swpRootFolder), Encoding.UTF8.GetBytes(json));
        }
    }
}
