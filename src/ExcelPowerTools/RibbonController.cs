using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using ExcelPowerTools.Commands;

namespace ExcelPowerTools
{
    /// <summary>
    /// Builds the "Power Tools" ribbon tab and routes button clicks to the command methods.
    /// Excel-DNA discovers this class (it derives from ExcelRibbon) and loads the XML automatically.
    /// </summary>
    [ComVisible(true)]
    public sealed class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='ptTab' label='Power Tools'>
        <group id='ptFields' label='Fields'>
          <button id='ptChangeField' label='Change Field...' size='large'
                  imageMso='TableColumnsInsertLeft' onAction='OnChangeField'
                  screentip='Edit a delimited field inside each selected cell'/>
        </group>
        <group id='ptSelection' label='Selection'>
          <button id='ptKeepValues' label='Keep Values Only' size='large'
                  imageMso='PasteValues' onAction='OnKeepValues'
                  screentip='Replace formulas with their values (visible cells only). Shortcut: Ctrl+Q'/>
          <button id='ptRegexReplace' label='Regex Replace...' size='large'
                  imageMso='ReplaceDialog' onAction='OnRegexReplace'
                  screentip='Regex find &amp; replace with capture groups ($1) and a before/after preview — scoped to the selection or the whole sheet'/>
          <button id='ptSmartFill' label='Smart Fill Down' imageMso='FillDown' onAction='OnSmartFill'
                  screentip='Fill down incrementing the trailing number (Z07→Z08…, IP last octet); two seeds set the step; skips hidden rows. Shortcut: Ctrl+Shift+D'/>
          <button id='ptUndo' label='Undo PT' imageMso='Undo' onAction='OnUndoPt'
                  screentip='Undo the last timeline step (50 deep). Plain Ctrl+Z also works right after a PT operation. Shortcut: Ctrl+Alt+Z'/>
          <button id='ptRedo' label='Redo PT' imageMso='Redo' onAction='OnRedoPt'
                  screentip='Redo an undone timeline step. Shortcut: Ctrl+Alt+Y'/>
          <button id='ptTimeline' label='Timeline' imageMso='History' onAction='OnTimeline'
                  screentip='Show the undo timeline (PT operations + tracked manual edits) and jump to any step'/>
          <toggleButton id='ptTrackEdits' label='Track edits' imageMso='TrackChanges'
                  onAction='OnToggleTrackEdits' getPressed='GetTrackEditsPressed'
                  screentip='Feed your manual cell edits into the undo timeline (typing, delete, paste-over-selection). Formatting and structural changes are not tracked.'/>
        </group>
        <group id='ptCsv' label='CSV'>
          <button id='ptCsvOpen' label='Open CSV/TXT' size='large' imageMso='FileOpen' onAction='OnOpenCsv'
                  screentip='Open a CSV/TXT without letting Excel corrupt it: cells load as Text, delimiter/encoding/line-ending detected, header/zebra styling applied'/>
          <button id='ptCsvSave' label='Save CSV' imageMso='Save' onAction='OnSaveCsv'
                  screentip='Save back to the opened file, byte-faithfully, with a .bak backup'/>
          <button id='ptCsvExport' label='Export as CSV' imageMso='SaveAs' onAction='OnExportCsv'
                  screentip='Write the active sheet to a CSV with a chosen separator'/>
          <button id='ptCsvRestyle' label='Re-style' imageMso='FormatPainter' onAction='OnRestyleCsv'
                  screentip='Re-detect the header and re-apply the CSV view styling'/>
          <button id='ptCsvBackup' label='Backup' imageMso='Copy' onAction='OnBackupCsv'
                  screentip='Copy the opened CSV file to a timestamped backup'/>
          <toggleButton id='ptCsvIntercept' label='Intercept opens' imageMso='TextToColumns'
                  onAction='OnToggleIntercept' getPressed='GetInterceptPressed'
                  screentip='When Excel opens a .csv natively (double-click, File Open), re-route it through the safe CSV pipeline'/>
        </group>
        <group id='ptData' label='Data'>
          <button id='ptStats' label='Column Stats' size='large' imageMso='AutoSum' onAction='OnColumnStats'
                  screentip='Live profile of the active cell&apos;s column: counts, distinct, numeric range, top values'/>
          <button id='ptSortAsc' label='Sort A-Z (nat)' imageMso='SortAscendingExcel' onAction='OnSortAsc'
                  screentip='Stable natural sort of the data block by the active column (DI2 before DI10). Undo restores the original order.'/>
          <button id='ptSortDesc' label='Sort Z-A (nat)' imageMso='SortDescendingExcel' onAction='OnSortDesc'
                  screentip='Stable natural sort, descending'/>
          <button id='ptDedup' label='Dedup by key' imageMso='RemoveDuplicates' onAction='OnDedup'
                  screentip='Remove duplicate rows keyed on the selected column(s), keeping the first occurrence. Undoable.'/>
        </group>
        <group id='ptWorkbook' label='Workbook'>
          <button id='ptSplitWb' label='Split into files...' imageMso='CreateReportFromWizard'
                  onAction='OnSplitWorkbook'
                  screentip='Split this workbook into several files of N sheets each'/>
        </group>
        <group id='ptFramework' label='Framework'>
          <button id='ptOpenVs' label='Open in VS' size='large' imageMso='ViewCode'
                  onAction='OnOpenInVs'
                  screentip='Open this add-in&apos;s solution in Visual Studio (the Alt+F11 equivalent)'/>
          <button id='ptOpenRepo' label='Repo folder' imageMso='Folder'
                  onAction='OnOpenRepo' screentip='Open the repository folder in Explorer'/>
          <button id='ptOpenOut' label='Build output' imageMso='FileOpen'
                  onAction='OnOpenBuildOutput' screentip='Open the folder the loaded .xll was built into'/>
          <button id='ptAbout' label='About' imageMso='Info'
                  onAction='OnAbout' screentip='Version, loaded .xll path, and resolved repo/solution'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // ---- ribbon callbacks -----------------------------------------------------------------
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
        public void OnToggleTrackEdits(IRibbonControl control, bool pressed) => Interop.ManualEditTracker.SetEnabled(pressed);
        public bool GetTrackEditsPressed(IRibbonControl control) => Interop.ManualEditTracker.Enabled;
        public void OnOpenCsv(IRibbonControl control) => CsvCommands.OpenCsv();
        public void OnSaveCsv(IRibbonControl control) => CsvCommands.SaveCsv();
        public void OnExportCsv(IRibbonControl control) => CsvCommands.ExportSheetAsCsv();
        public void OnRestyleCsv(IRibbonControl control) => CsvCommands.RestyleActiveSheet();
        public void OnBackupCsv(IRibbonControl control) => CsvCommands.BackupActiveFile();
        public void OnToggleIntercept(IRibbonControl control, bool pressed) => CsvCommands.SetInterceptOpen(pressed);
        public bool GetInterceptPressed(IRibbonControl control) => CsvCommands.InterceptOpenEnabled();
        public void OnSplitWorkbook(IRibbonControl control) => WorkbookCommands.ExportSheetsAsFiles();

        // Framework / dev-loop callbacks
        public void OnOpenInVs(IRibbonControl control) => DevCommands.OpenSolutionInVisualStudio();
        public void OnOpenRepo(IRibbonControl control) => DevCommands.OpenRepoFolder();
        public void OnOpenBuildOutput(IRibbonControl control) => DevCommands.OpenBuildOutputFolder();
        public void OnAbout(IRibbonControl control) => DevCommands.About();
    }
}
