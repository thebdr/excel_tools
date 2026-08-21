using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Integration.CustomUI;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Interop;
using ExcelPowerTools.UI;

namespace ExcelPowerTools
{
    /// <summary>
    /// The ribbon is deliberately a compact command strip. Less frequent tools, configuration and
    /// experimental features live in the searchable Power Tools Hub task pane.
    /// </summary>
    [ComVisible(true)]
    public sealed class RibbonController : ExcelRibbon
    {
#if DEBUG
        internal const string RibbonTabLabel = "Power Tools D";
#else
        internal const string RibbonTabLabel = "Power Tools";
#endif

        private IRibbonUI _ribbon;
        private bool _eventsAttached;

        public override string GetCustomUI(string ribbonId)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'
          onLoad='OnRibbonLoad' loadImage='LoadImage'>
  <ribbon>
    <tabs>
      <tab id='ptTab' label='" + RibbonTabLabel + @"'>
        <group id='ptHome' label='Power Tools'>
          <toggleButton id='ptHub' label='Tools Hub' size='large' image='ptHubIcon'
                  onAction='OnToggleHub' getPressed='GetHubPressed'
                  screentip='Open the Power Tools Hub'
                  supertip='Browse and search every command, change add-in settings, and reach specialist or experimental tools in one docked window.'/>
          <button id='ptUndo' label='Undo PT' imageMso='Undo' onAction='OnUndoPt'
                  getEnabled='GetUndoEnabled'
                  screentip='Undo the latest Power Tools timeline step (Ctrl+Alt+Z)'/>
          <button id='ptRedo' label='Redo PT' imageMso='Redo' onAction='OnRedoPt'
                  getEnabled='GetRedoEnabled'
                  screentip='Redo the latest undone timeline step (Ctrl+Alt+Y)'/>
          <button id='ptTimeline' label='Timeline' image='ptTimelineIcon16' onAction='OnTimeline'
                  screentip='Inspect the full undo and manual-edit timeline'/>
        </group>

        <group id='ptCompare' label='Compare &amp; review'>
          <splitButton id='ptCompareSplit' size='large'>
            <button id='ptComparePrimary' label='Compare revisions' image='ptCompareIcon' onAction='OnCompareRevisions'
                    screentip='Build a multi-revision comparison project'
                    supertip='Choose every available revision, correct the inferred order and parent links, then build an integrity-checked Review workbook with aligned logical changes and history.'/>
            <menu id='ptCompareMenu' label='Compare revisions' itemSize='large'>
              <button id='ptCompareProjectMenu' label='New comparison project...' image='ptCompareIcon' onAction='OnCompareRevisions'
                      screentip='Compare two to twenty revisions with editable lineage and conservative alignment'/>
              <button id='ptCompareFilesMenu' label='Exact two-file preview...' image='ptCompareIcon' onAction='OnCompareFiles'
                      screentip='Generate a positional exact pairwise Review workbook'/>
              <button id='ptOpenReviewMenu' label='Review workbook in Hub...' imageMso='FileOpen' onAction='OnOpenReview'
                      screentip='Validate a saved Review workbook and continue its decision ledger in the Hub'/>
            </menu>
          </splitButton>
        </group>

        <group id='ptDelimited' label='Delimited workspace'>
          <splitButton id='ptOpenSplit' size='large'>
            <button id='ptOpenPrimary' label='Open data' image='ptOpenIcon' onAction='OnOpenCsv'
                    screentip='Open delimited files'
                    supertip='Open one or more CSV, TSV or separated-value TXT files as sheets in one indexed workspace.'/>
            <menu id='ptOpenMenu' label='Open data' itemSize='large'>
              <button id='ptOpenFilesMenu' label='Open files...' imageMso='FileOpen' onAction='OnOpenCsv'
                      screentip='Choose one or more CSV, TSV or separated-value TXT files'/>
              <button id='ptOpenFolderMenu' label='Open whole folder...' imageMso='FileOpen' onAction='OnOpenCsvFolder'
                      screentip='Open every supported delimited file in a folder'/>
            </menu>
          </splitButton>
          <button id='ptCsvSave' label='Save file' size='large' image='ptSaveIcon' onAction='OnSaveCsv'
                  getEnabled='GetCsvSourceEnabled'
                  screentip='Safely save the active source file'
                  supertip='Write the active source sheet back to disk with external-change protection and a backup.'/>
          <button id='ptCsvSaveAll' label='Save all' image='ptSaveAllIcon16' onAction='OnSaveAllCsv'
                  getEnabled='GetCsvWorkspaceEnabled'
                  screentip='Save every edited source sheet in this workspace'/>
          <button id='ptCsvIndex' label='Index' image='ptIndexIcon16' onAction='OnCsvIndex'
                  getEnabled='GetCsvWorkspaceEnabled'
                  screentip='Refresh and show the workspace file index'/>
        </group>

        <group id='ptTransform' label='Transform'>
          <button id='ptKeepValues' label='Keep values' size='large' image='ptValuesIcon' onAction='OnKeepValues'
                  screentip='Replace formulas with their current values'
                  supertip='Convert formulas to values in visible selected cells. Shortcut: Ctrl+Q.'/>
          <button id='ptSmartFill' label='Smart fill' imageMso='FillDown' onAction='OnSmartFill'
                  screentip='Fill down intelligently (Ctrl+Shift+D)'/>
          <button id='ptChangeField' label='Change field...' imageMso='TableColumnsInsertLeft' onAction='OnChangeField'
                  screentip='Edit one delimited field inside every selected cell'/>
          <button id='ptRegexReplace' label='Regex replace...' imageMso='ReplaceDialog' onAction='OnRegexReplace'
                  screentip='Preview and apply regex replacements'/>
        </group>

        <group id='ptIoBridge' label='IO Bridge'>
          <button id='ptIobOpen' label='Open workspace' size='large' imageMso='FileOpen' onAction='OnIobOpen'
                  screentip='Build the IO Bridge workspace'
                  supertip='Join an IO-list workbook with its Plant Inputs/Outputs CSVs. Edits queue instead of applying; a batch apply writes the real files.'/>
          <button id='ptIobRowActions' label='Row actions' imageMso='ListBullets' onAction='OnIobRowActions'
                  screentip='Queue linked-field updates, integrate, disable/enable or delete for the selected row'/>
          <button id='ptIobApply' label='Apply queue' imageMso='ReviewAcceptChange' onAction='OnIobApply'
                  screentip='Write the included pending actions to the CSVs and the IO-list workbook'/>
          <button id='ptIobRefresh' label='Refresh' imageMso='RefreshAll' onAction='OnIobRefresh'
                  screentip='Rebuild the bridge sheets from the current documents; the queue re-anchors'/>
        </group>

        <group id='ptSwpBridge' label='SWP Bridge'>
          <button id='ptSwpOpen' label='Open workspace' size='large' imageMso='FileOpen' onAction='OnSwpOpen'
                  screentip='Build the SWP Bridge workspace'
                  supertip='Join the IO list&#39;s diagnostic mnemonics to the SWP workbook as a reviewable diff; a headless batch apply patches the SWP.'/>
          <button id='ptSwpApply' label='Apply queue' imageMso='ReviewAcceptChange' onAction='OnSwpApply'
                  screentip='Write the included pending actions into the SWP (in place or into a new dated version folder)'/>
          <button id='ptSwpRefresh' label='Refresh' imageMso='RefreshAll' onAction='OnSwpRefresh'
                  screentip='Re-read the documents, recompute the diff, re-render the workspace'/>
          <button id='ptSwpDashboard' label='Dashboard' imageMso='ChartPlacement' onAction='OnSwpDashboard'
                  screentip='Show the SWP Bridge dashboard (progress, tallies, findings)'/>
        </group>

        <group id='ptData' label='Clean &amp; analyze'>
          <splitButton id='ptSortSplit'>
            <button id='ptSortPrimary' label='Natural sort' imageMso='SortAscendingExcel' onAction='OnSortAsc'
                    screentip='Natural sort ascending (DI2 before DI10)'/>
            <menu id='ptSortMenu' label='Natural sort' itemSize='normal'>
              <button id='ptSortAscMenu' label='Ascending' imageMso='SortAscendingExcel' onAction='OnSortAsc'/>
              <button id='ptSortDescMenu' label='Descending' imageMso='SortDescendingExcel' onAction='OnSortDesc'/>
            </menu>
          </splitButton>
          <button id='ptDedup' label='Deduplicate' imageMso='RemoveDuplicates' onAction='OnDedup'
                  screentip='Remove duplicate rows by the selected key columns'/>
          <button id='ptStats' label='Column stats' imageMso='AutoSum' onAction='OnColumnStats'
                  screentip='Open a live profile for the active column'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public override object LoadImage(string imageId)
            => RibbonIcons.Get(imageId) ?? base.LoadImage(imageId);

        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            if (_eventsAttached) return;

            PowerToolsPaneManager.VisibilityChanged += OnPaneVisibilityChanged;
            UndoManager.Changed += OnHistoryChanged;
            SelectionEvents.Changed += OnExcelContextChanged;
            _eventsAttached = true;
        }

        public override void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (_eventsAttached)
            {
                PowerToolsPaneManager.VisibilityChanged -= OnPaneVisibilityChanged;
                UndoManager.Changed -= OnHistoryChanged;
                SelectionEvents.Changed -= OnExcelContextChanged;
                _eventsAttached = false;
            }
            _ribbon = null;
            base.OnDisconnection(removeMode, ref custom);
        }

        private void OnPaneVisibilityChanged() => InvalidateControl("ptHub");

        private void OnHistoryChanged()
        {
            InvalidateControl("ptUndo");
            InvalidateControl("ptRedo");
        }

        private void OnExcelContextChanged()
        {
            InvalidateControl("ptCsvSave");
            InvalidateControl("ptCsvSaveAll");
            InvalidateControl("ptCsvIndex");
        }

        private void InvalidateControl(string id)
        {
            try { _ribbon?.InvalidateControl(id); }
            catch { }
        }

        // ---- primary ribbon callbacks -----------------------------------------------------------

        public void OnToggleHub(IRibbonControl control, bool pressed) => PowerToolsPaneManager.SetVisible(pressed);
        public bool GetHubPressed(IRibbonControl control) => PowerToolsPaneManager.IsVisible;
        public bool GetUndoEnabled(IRibbonControl control) => UndoManager.UndoEntries.Count > 0;
        public bool GetRedoEnabled(IRibbonControl control) => UndoManager.RedoEntries.Count > 0;
        public bool GetCsvSourceEnabled(IRibbonControl control)
        {
            try
            {
                dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
                return CsvSession.Get(app?.ActiveSheet) != null;
            }
            catch { return false; }
        }

        public bool GetCsvWorkspaceEnabled(IRibbonControl control)
        {
            try
            {
                dynamic app = ExcelDna.Integration.ExcelDnaUtil.Application;
                return CsvSession.IsWorkspaceWorkbook(app?.ActiveWorkbook);
            }
            catch { return false; }
        }

        public void OnChangeField(IRibbonControl control) => FieldEditorCommands.ChangeField();
        public void OnKeepValues(IRibbonControl control) => SelectionCommands.KeepValuesOnly();
        public void OnRegexReplace(IRibbonControl control) => RegexReplaceCommands.ShowRegexReplace();
        public void OnSmartFill(IRibbonControl control) => FillCommands.SmartFillDown();
        public void OnUndoPt(IRibbonControl control) => UndoCommands.UndoLast();
        public void OnRedoPt(IRibbonControl control) => UndoCommands.RedoLast();
        public void OnTimeline(IRibbonControl control) => UndoCommands.ShowTimeline();
        public void OnColumnStats(IRibbonControl control) => DataCommands.ShowColumnStats();
        public void OnSortAsc(IRibbonControl control) => DataCommands.SortAscending();
        public void OnSortDesc(IRibbonControl control) => DataCommands.SortDescending();
        public void OnDedup(IRibbonControl control) => DataCommands.DedupRows();
        public void OnOpenCsv(IRibbonControl control) => CsvCommands.OpenCsv();
        public void OnOpenCsvFolder(IRibbonControl control) => CsvCommands.OpenCsvFolder();
        public void OnSaveCsv(IRibbonControl control) => CsvCommands.SaveCsv();
        public void OnSaveAllCsv(IRibbonControl control) => CsvCommands.SaveAllCsv();
        public void OnCsvIndex(IRibbonControl control) => CsvCommands.ShowIndex();
        public void OnCompareRevisions(IRibbonControl control) => ComparisonProjectCommands.StartComparisonProject();
        public void OnCompareFiles(IRibbonControl control) => ComparisonCommands.StartExactPreview();
        public void OnOpenReview(IRibbonControl control) => PowerToolsPaneManager.OpenComparisonReviewPicker();

        // ---- callbacks retained for the hub and macro compatibility ------------------------------

        public void OnExportCsv(IRibbonControl control) => CsvCommands.ExportSheetAsCsv();
        public void OnRestyleCsv(IRibbonControl control) => CsvCommands.RestyleActiveSheet();
        public void OnBackupCsv(IRibbonControl control) => CsvCommands.BackupActiveFile();
        public void OnToggleIntercept(IRibbonControl control, bool pressed) => CsvCommands.SetInterceptOpen(pressed);
        public bool GetInterceptPressed(IRibbonControl control) => CsvCommands.InterceptOpenEnabled();
        public void OnToggleTrackEdits(IRibbonControl control, bool pressed) => ManualEditTracker.SetEnabled(pressed);
        public bool GetTrackEditsPressed(IRibbonControl control) => ManualEditTracker.Enabled;
        public void OnSplitWorkbook(IRibbonControl control) => WorkbookCommands.ExportSheetsAsFiles();
        public void OnXBuildAddress(IRibbonControl control) => IoPipeline.IoListCommands.BuildAddressColumn();
        public void OnXPullIoData(IRibbonControl control) => IoPipeline.IoListCommands.PullFromIoList();
        public void OnXPullMnemonics(IRibbonControl control) => IoPipeline.IoListCommands.PullMnemonicsIntoIoList();

        public void OnIobOpen(IRibbonControl control) => IoPipeline.IoBridge.IoBridgeCommands.OpenWorkspace();
        public void OnIobRowActions(IRibbonControl control) => IoPipeline.IoBridge.IoBridgeCommands.RowActions();
        public void OnIobApply(IRibbonControl control) => IoPipeline.IoBridge.IoBridgeCommands.ApplyBatch();
        public void OnIobRefresh(IRibbonControl control) => IoPipeline.IoBridge.IoBridgeCommands.RefreshWorkspace();

        public void OnSwpOpen(IRibbonControl control) => IoPipeline.SwpBridge.SwpBridgeCommands.OpenWorkspace();
        public void OnSwpApply(IRibbonControl control) => IoPipeline.SwpBridge.SwpBridgeCommands.ApplyBatch();
        public void OnSwpRefresh(IRibbonControl control) => IoPipeline.SwpBridge.SwpBridgeCommands.RefreshWorkspace();
        public void OnSwpDashboard(IRibbonControl control) => IoPipeline.SwpBridge.SwpBridgeCommands.ShowDashboard();
        public void OnOpenInVs(IRibbonControl control) => DevCommands.OpenSolutionInVisualStudio();
        public void OnOpenRepo(IRibbonControl control) => DevCommands.OpenRepoFolder();
        public void OnOpenBuildOutput(IRibbonControl control) => DevCommands.OpenBuildOutputFolder();
        public void OnAbout(IRibbonControl control) => DevCommands.About();
    }
}
