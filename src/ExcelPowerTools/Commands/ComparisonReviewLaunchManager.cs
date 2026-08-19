using System;
using System.Collections.Generic;
using System.IO;
using ExcelDna.Integration;
using ExcelPowerTools.Core.Comparison.Persistence;
using ExcelPowerTools.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Retains each review launch copy's protective read handle for exactly as long as Excel still has that
    /// private workbook open. It never closes an existing user workbook and never deletes a source report.
    /// </summary>
    internal static class ComparisonReviewLaunchManager
    {
        private static readonly Dictionary<string, ReviewWorkbookLaunchCopy> Launches =
            new Dictionary<string, ReviewWorkbookLaunchCopy>(StringComparer.OrdinalIgnoreCase);

        internal static void Register(ReviewWorkbookLaunchCopy launchCopy)
        {
            if (launchCopy == null) throw new ArgumentNullException(nameof(launchCopy));
            if (Launches.ContainsKey(launchCopy.LaunchPath))
                throw new InvalidOperationException("The private review launch copy is already registered.");
            Launches.Add(launchCopy.LaunchPath, launchCopy);
        }

        internal static void OnWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            string path = TryGetWorkbookPath(workbook);
            if (path != null && Launches.ContainsKey(path)) QueueSweep();
        }

        internal static void OnWorkbookAfterSave(Excel.Workbook workbook, bool success)
        {
            // Save As changes FullName and releases the old read-only launch path. Sweep after Excel finishes
            // the save so the user-created destination remains completely outside our cleanup ownership.
            QueueSweep();
        }

        internal static void Shutdown()
        {
            foreach (ReviewWorkbookLaunchCopy launch in Launches.Values)
            {
                try { launch.Dispose(); }
                catch { }
            }
            Launches.Clear();
        }

        private static void QueueSweep()
        {
            try { ExcelAsyncUtil.QueueAsMacro(SweepClosedLaunches); }
            catch { }
        }

        private static void SweepClosedLaunches()
        {
            var livePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                dynamic app = ExcelApp.App;
                int count = Convert.ToInt32(app.Workbooks.Count);
                for (int index = 1; index <= count; index++)
                {
                    string path = TryGetWorkbookPath(app.Workbooks[index]);
                    if (path != null) livePaths.Add(path);
                }
            }
            catch
            {
                // An uncertain Excel enumeration must retain the protective handles. A later close/save event
                // or AutoClose will retry; premature cleanup would make a live launch path mutable.
                return;
            }

            var retired = new List<string>();
            foreach (KeyValuePair<string, ReviewWorkbookLaunchCopy> pair in Launches)
            {
                if (livePaths.Contains(pair.Key)) continue;
                try { pair.Value.Dispose(); }
                catch { }
                retired.Add(pair.Key);
            }
            foreach (string path in retired) Launches.Remove(path);
        }

        private static string TryGetWorkbookPath(dynamic workbook)
        {
            try
            {
                string fullName = (string)workbook.FullName;
                return string.IsNullOrWhiteSpace(fullName) ? null : Path.GetFullPath(fullName);
            }
            catch { return null; }
        }
    }
}
