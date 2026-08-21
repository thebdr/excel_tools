using System;
using System.Text;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using ExcelPowerTools.Core.Files;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    /// <summary>
    /// The queue's disk home: a JSON sidecar NEXT TO the IO-list workbook, written atomically on every
    /// queue mutation so an Excel crash loses nothing beyond the last keystroke.
    /// </summary>
    internal static class IoBridgeSidecarStore
    {
        public static string PathFor(string ioListPath) => ioListPath + ".iobridge.json";

        /// <summary>
        /// Loads the sidecar for an IO list. Returns null when none exists; throws with a recovery
        /// hint when one exists but cannot be read — silently dropping a queue is never acceptable.
        /// </summary>
        public static IoBridgeSidecar Load(string ioListPath)
        {
            string path = PathFor(ioListPath);
            SourceReadState state = AtomicFileStore.ReadSource(path, out byte[] bytes, out string readError);
            if (state == SourceReadState.Missing) return null;
            if (state == SourceReadState.Unavailable)
                throw new InvalidOperationException("The queue sidecar could not be read:\n" + path + "\n" + readError);

            string json = Encoding.UTF8.GetString(bytes);
            if (json.Length > 0 && json[0] == (char)0xFEFF) json = json.Substring(1);   // tolerate a UTF-8 BOM

            if (!IoBridgeQueueJson.TryParse(json, out IoBridgeSidecar sidecar, out string parseError))
                throw new InvalidOperationException(
                    "The queue sidecar is damaged and was not loaded:\n" + path + "\n" + parseError +
                    "\n\nA previous version may exist as \"" + System.IO.Path.GetFileName(path) + ".bak\".");
            return sidecar;
        }

        public static void Save(IoBridgeSession session)
        {
            var sidecar = new IoBridgeSidecar
            {
                GeneratedUtc = DateTime.UtcNow.ToString("O"),
                IoListPath = session.IoListPath,
                InputsCsvPath = session.InputsCsvPath,
                OutputsCsvPath = session.OutputsCsvPath,
                Queue = session.Queue,
            };
            sidecar.CardSheets.AddRange(session.CardSheets);

            string json = IoBridgeQueueJson.Serialize(sidecar);
            AtomicFileStore.WriteWithBackup(session.SidecarPath, Encoding.UTF8.GetBytes(json));
        }
    }
}
