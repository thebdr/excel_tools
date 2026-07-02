using ExcelDna.Integration;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Undo/redo for Power-Tools operations. Native Ctrl+Z works as a one-shot right after an operation
    /// (via Application.OnUndo); these commands reach the full 20-deep stack at any time.
    /// </summary>
    public static class UndoCommands
    {
        [ExcelCommand(Name = "PT_UndoLast", MenuName = "Power Tools", MenuText = "Undo Power Tools operation", ShortCut = "^%z")]
        public static void UndoLast() => UndoManager.UndoLast();

        [ExcelCommand(Name = "PT_RedoLast", MenuName = "Power Tools", MenuText = "Redo Power Tools operation", ShortCut = "^%y")]
        public static void RedoLast() => UndoManager.RedoLast();
    }
}
