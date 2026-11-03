using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Core.Comparison.RevisionBridge;
using ExcelPowerTools.Interop;
using ExcelPowerTools.UI;

namespace ExcelPowerTools.Commands
{
    public static class RevisionComparisonCommands
    {
        private const string WorkbookFilter =
            "Excel workbooks|*.xlsx;*.xlsm|Excel workbook (*.xlsx)|*.xlsx|Macro-enabled workbook (*.xlsm)|*.xlsm";

        [ExcelCommand(
            Name = "PT_CompareTableRevisions",
            Description = "Compare ordered table revisions by explicit keys",
            MenuText = "Compare table revisions...")]
        public static void StartComparison()
        {
            try
            {
                string[] paths;
                using (var picker = new RevisionComparisonFilePickerForm(WorkbookFilter))
                {
                    if (picker.ShowDialog(ExcelWindow.Current) != DialogResult.OK) return;
                    paths = picker.SelectedPaths.ToArray();
                }

                Cursor previous = Cursor.Current;
                RevisionComparisonInspection inspection;
                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    inspection = RevisionComparisonEngine.Inspect(paths);
                }
                finally
                {
                    Cursor.Current = previous;
                }

                RevisionComparisonPlan plan;
                using (var setup = new RevisionComparisonSetupForm(inspection))
                {
                    if (setup.ShowDialog(ExcelWindow.Current) != DialogResult.OK) return;
                    plan = setup.ResultPlan;
                }

                RevisionComparisonResult result;
                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    result = RevisionComparisonEngine.Compare(inspection, plan);
                }
                finally
                {
                    Cursor.Current = previous;
                }

                string reviewPath;
                using (var save = new SaveFileDialog
                {
                    Title = "Save comparison review",
                    Filter = "Excel workbook (*.xlsx)|*.xlsx",
                    DefaultExt = "xlsx",
                    AddExtension = true,
                    OverwritePrompt = true,
                    InitialDirectory = System.IO.Path.GetDirectoryName(paths[0]),
                    FileName = "Table Comparison Review.xlsx",
                })
                {
                    if (save.ShowDialog(ExcelWindow.Current) != DialogResult.OK) return;
                    reviewPath = save.FileName;
                }

                RevisionComparisonReviewFile.Write(reviewPath, result);
                string openWarning = TryOpenWorkbook(reviewPath);
                Dialogs.Info(
                    "Comparison review created\n\n" + reviewPath + "\n\n" +
                    result.ChangedCellCount.ToString(CultureInfo.InvariantCulture) + " changed cells, " +
                    result.AddedRowCount.ToString(CultureInfo.InvariantCulture) + " added rows, " +
                    result.RemovedRows.Count.ToString(CultureInfo.InvariantCulture) + " removed rows.\n\n" +
                    "Type Yes in Apply on Added or Changed, save the review, then choose Apply selected comparison changes." +
                    OpenWarning(openWarning));
            }
            catch (Exception exception)
            {
                Dialogs.Info("Table comparison could not be completed\n\n" + ComparisonCommands.DescribeException(exception));
            }
        }

        [ExcelCommand(
            Name = "PT_ApplyTableRevisionChanges",
            Description = "Apply selected keyed comparison actions to a new workbook",
            MenuText = "Apply selected comparison changes...")]
        public static void ApplySelected()
        {
            try
            {
                string reviewPath = ActiveReviewPath();
                if (reviewPath == null)
                {
                    using (var picker = new OpenFileDialog
                    {
                        Title = "Open a table comparison review",
                        Filter = "Excel workbook (*.xlsx)|*.xlsx",
                        Multiselect = false,
                        CheckFileExists = true,
                    })
                    {
                        if (picker.ShowDialog(ExcelWindow.Current) != DialogResult.OK) return;
                        reviewPath = picker.FileName;
                    }
                }

                RevisionComparisonReviewSelection selection = RevisionComparisonReviewFile.ReadSelection(reviewPath);
                if (selection.SelectedActionIds.Count == 0)
                {
                    Dialogs.Info(
                        "No actions are selected. Type Yes in the Apply column on Added or Changed, save the review, then retry.");
                    return;
                }
                string targetPath = selection.Sources[selection.Plan.ApplyTargetRevisionIndex].Path;
                string extension = System.IO.Path.GetExtension(targetPath);
                string outputPath;
                using (var save = new SaveFileDialog
                {
                    Title = "Create a new workbook with selected changes",
                    Filter = string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase)
                        ? "Macro-enabled workbook (*.xlsm)|*.xlsm"
                        : "Excel workbook (*.xlsx)|*.xlsx",
                    DefaultExt = extension.TrimStart('.'),
                    AddExtension = true,
                    OverwritePrompt = true,
                    InitialDirectory = System.IO.Path.GetDirectoryName(targetPath),
                    FileName = System.IO.Path.GetFileNameWithoutExtension(targetPath) + " - compared" + extension,
                })
                {
                    if (save.ShowDialog(ExcelWindow.Current) != DialogResult.OK) return;
                    outputPath = save.FileName;
                }
                if (File.Exists(outputPath))
                {
                    Dialogs.Info("Choose a new filename. Existing workbooks are never replaced by comparison apply.");
                    return;
                }

                Cursor previous = Cursor.Current;
                RevisionComparisonApplyResult result;
                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    result = RevisionComparisonApplyEngine.ApplyReview(reviewPath, outputPath);
                }
                finally
                {
                    Cursor.Current = previous;
                }

                string openWarning = TryOpenWorkbook(result.OutputPath);
                Dialogs.Info(
                    "Verified output created\n\n" + result.OutputPath + "\n\n" +
                    result.ChangedCellCount.ToString(CultureInfo.InvariantCulture) + " cell changes applied, " +
                    result.AddedRowCount.ToString(CultureInfo.InvariantCulture) + " rows inserted, " +
                    result.AlreadyCurrentCount.ToString(CultureInfo.InvariantCulture) + " actions already current.\n\n" +
                    "Every source revision remains unchanged." + OpenWarning(openWarning));
            }
            catch (Exception exception)
            {
                Dialogs.Info("Selected changes were not applied\n\n" + ComparisonCommands.DescribeException(exception));
            }
        }

        private static string ActiveReviewPath()
        {
            dynamic workbook;
            string path;
            try
            {
                dynamic application = ExcelDnaUtil.Application;
                workbook = application.ActiveWorkbook;
                if (workbook == null) return null;
                path = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
            }
            catch
            {
                return null;
            }

            try
            {
                RevisionComparisonReviewFile.ReadSelection(path);
            }
            catch
            {
                return null;
            }

            workbook.Save();
            return System.IO.Path.GetFullPath(path);
        }

        private static string TryOpenWorkbook(string path)
        {
            try
            {
                OpenWorkbook(path, readOnly: false);
                return null;
            }
            catch (Exception exception)
            {
                return ComparisonCommands.DescribeException(exception);
            }
        }

        private static string OpenWarning(string warning) =>
            string.IsNullOrWhiteSpace(warning)
                ? ""
                : "\n\nThe file was created and verified, but Excel could not open it automatically:\n" + warning;

        private static void OpenWorkbook(string path, bool readOnly)
        {
            dynamic application = ExcelDnaUtil.Application;
            foreach (dynamic workbook in application.Workbooks)
            {
                string openPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                if (!string.Equals(openPath, path, StringComparison.OrdinalIgnoreCase)) continue;
                workbook.Activate();
                return;
            }
            application.Workbooks.Open(path, 0, readOnly);
        }
    }
}
