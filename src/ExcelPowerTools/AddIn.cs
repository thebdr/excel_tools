using ExcelDna.Integration;
using ExcelPowerTools.Commands;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPowerTools
{
    /// <summary>
    /// Add-in lifecycle: subscribes to the application-level WorkbookOpen event so a natively-opened
    /// .csv can be re-routed through the corruption-safe CSV pipeline.
    /// </summary>
    public sealed class AddIn : IExcelAddIn
    {
        // Held for the lifetime of the add-in so the COM event sink is not garbage-collected.
        private static Excel.Application _app;

        public void AutoOpen()
        {
            _app = (Excel.Application)ExcelDnaUtil.Application;
            _app.WorkbookOpen += CsvCommands.OnWorkbookOpen;
            _app.WorkbookBeforeClose += CsvCommands.OnWorkbookBeforeClose;
            _app.WorkbookBeforeClose += ComparisonReviewLaunchManager.OnWorkbookBeforeClose;
            _app.WorkbookAfterSave += ComparisonReviewLaunchManager.OnWorkbookAfterSave;
            _app.WorkbookBeforeClose += OnWorkbookBeforeCloseHub;
            _app.SheetActivate += CsvCommands.OnSheetActivate;
            _app.SheetBeforeDelete += CsvCommands.OnSheetBeforeDelete;

            // Manual-edit tracking for the unified undo timeline.
            _app.SheetSelectionChange += Interop.ManualEditTracker.OnSelectionChange;
            _app.SheetChange += Interop.ManualEditTracker.OnSheetChange;
            _app.SheetActivate += Interop.ManualEditTracker.OnSheetActivate;

            // Runs after manual tracking so any exceptional CSV sentinel move can join the same undo step.
            _app.SheetChange += CsvCommands.OnSheetChange;

            // IO Bridge edit capture — bridge sheets are excluded from manual tracking, so order is free.
            _app.SheetChange += Experimental.IoBridge.IoBridgeEditCapture.OnSheetChange;
            _app.SheetBeforeDoubleClick += Experimental.IoBridge.IoBridgeEditCapture.OnSheetBeforeDoubleClick;
            _app.WorkbookBeforeClose += Experimental.IoBridge.IoBridgeEditCapture.OnWorkbookBeforeClose;

            // Selection broadcast for live panels (Column Stats).
            _app.SheetSelectionChange += OnSelectionChangedHub;
            _app.SheetActivate += OnSheetActivateHub;
            _app.WindowActivate += OnWindowActivateHub;

            // Show a readable message instead of a bare #VALUE! if a UDF throws.
            ExcelIntegration.RegisterUnhandledExceptionHandler(
                ex => "!ERR: " + ((ex as System.Exception)?.Message ?? ex?.ToString()));
        }

        public void AutoClose()
        {
            UI.PowerToolsPaneManager.Shutdown();
            ComparisonReviewLaunchManager.Shutdown();
            if (_app == null) return;
            _app.WorkbookOpen -= CsvCommands.OnWorkbookOpen;
            _app.WorkbookBeforeClose -= CsvCommands.OnWorkbookBeforeClose;
            _app.WorkbookBeforeClose -= ComparisonReviewLaunchManager.OnWorkbookBeforeClose;
            _app.WorkbookAfterSave -= ComparisonReviewLaunchManager.OnWorkbookAfterSave;
            _app.WorkbookBeforeClose -= OnWorkbookBeforeCloseHub;
            _app.SheetChange -= CsvCommands.OnSheetChange;
            _app.SheetChange -= Experimental.IoBridge.IoBridgeEditCapture.OnSheetChange;
            _app.SheetBeforeDoubleClick -= Experimental.IoBridge.IoBridgeEditCapture.OnSheetBeforeDoubleClick;
            _app.WorkbookBeforeClose -= Experimental.IoBridge.IoBridgeEditCapture.OnWorkbookBeforeClose;
            _app.SheetActivate -= CsvCommands.OnSheetActivate;
            _app.SheetBeforeDelete -= CsvCommands.OnSheetBeforeDelete;
            _app.SheetSelectionChange -= Interop.ManualEditTracker.OnSelectionChange;
            _app.SheetChange -= Interop.ManualEditTracker.OnSheetChange;
            _app.SheetActivate -= Interop.ManualEditTracker.OnSheetActivate;
            _app.SheetSelectionChange -= OnSelectionChangedHub;
            _app.SheetActivate -= OnSheetActivateHub;
            _app.WindowActivate -= OnWindowActivateHub;
            _app = null;
        }

        private static void OnSelectionChangedHub(object sh, Excel.Range target) => Interop.SelectionEvents.Raise();
        private static void OnSheetActivateHub(object sh) => Interop.SelectionEvents.Raise();
        private static void OnWindowActivateHub(Excel.Workbook workbook, Excel.Window window)
        {
            UI.PowerToolsPaneManager.OnWindowActivated();
            Interop.SelectionEvents.Raise();
        }

        private static void OnWorkbookBeforeCloseHub(Excel.Workbook workbook, ref bool cancel)
            => UI.PowerToolsPaneManager.QueueWindowSweep();
    }
}
