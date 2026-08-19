using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Owns one reusable Hub pane per Excel window. Office task panes are window-scoped, so the ribbon
    /// toggle always reflects the pane belonging to the currently active window.
    /// </summary>
    internal static class PowerToolsPaneManager
    {
        private const int DefaultWidth = 410;
        private static readonly Dictionary<int, PaneEntry> Entries = new Dictionary<int, PaneEntry>();

        public static event Action VisibilityChanged;

        public static bool IsVisible
        {
            get
            {
                if (!TryGetActiveWindow(out int windowHandle, out _)) return false;
                if (!Entries.TryGetValue(windowHandle, out PaneEntry entry)) return false;
                try { return entry.Pane.Visible; }
                catch (Exception ex)
                {
                    if (IsDisconnected(ex)) RetireManagedEntry(windowHandle, entry);
                    return false;
                }
            }
        }

        public static void SetVisible(bool visible)
        {
            if (!TryGetActiveWindow(out int windowHandle, out object window))
            {
                RaiseVisibilityChanged();
                return;
            }

            Entries.TryGetValue(windowHandle, out PaneEntry entry);

            if (entry == null)
            {
                if (!visible)
                {
                    RaiseVisibilityChanged();
                    return;
                }
                entry = CreatePane(windowHandle, window);
                if (entry == null)
                {
                    RaiseVisibilityChanged();
                    return;
                }
            }

            try
            {
                if (visible) entry.Control.RefreshState();
                entry.Pane.Visible = visible;
            }
            catch (Exception ex)
            {
                if (IsDisconnected(ex))
                {
                    RetireManagedEntry(windowHandle, entry);
                    if (visible) TryCreateAndShow(windowHandle, window);
                }
                else
                    Dialogs.Info("Excel is temporarily busy; the Power Tools Hub could not be " +
                                 (visible ? "opened" : "closed") + ":\n" + ex.Message);
            }
            finally { RaiseVisibilityChanged(); }
        }

        /// <summary>Called by Excel's WindowActivate event so the ribbon toggle follows the active window.</summary>
        public static void OnWindowActivated()
        {
            SweepClosedWindows();
            if (TryGetActiveWindow(out int handle, out _) && Entries.TryGetValue(handle, out PaneEntry entry))
            {
                try { if (entry.Pane.Visible) entry.Control.RefreshState(); }
                catch (Exception ex) { if (IsDisconnected(ex)) RetireManagedEntry(handle, entry); }
            }
            RaiseVisibilityChanged();
        }

        /// <summary>Runs after WorkbookBeforeClose has finished, so cancelled closes remain in the live set.</summary>
        public static void QueueWindowSweep()
        {
            try
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    SweepClosedWindows();
                    SelectionEvents.Raise();
                });
            }
            catch { }
        }

        /// <summary>
        /// Excel-DNA owns deletion of its custom task panes during COM disconnection. At AutoClose we only
        /// release managed subscriptions and strong references; touching the already-released COM panes can
        /// keep Excel alive or throw an InvalidComObjectException.
        /// </summary>
        public static void Shutdown()
        {
            foreach (PaneEntry entry in Entries.Values)
            {
                try { entry.Control.Detach(); } catch { }
                try { entry.Control.Dispose(); } catch { }
            }
            Entries.Clear();
        }

        private static PaneEntry CreatePane(int windowHandle, object window)
        {
            var control = new PowerToolsPane();
            control.AttachWindow(windowHandle);
            CustomTaskPane pane = null;
            try
            {
                pane = CustomTaskPaneFactory.CreateCustomTaskPane(
                    control, "Power Tools Hub", window);
                if (pane == null) throw new InvalidOperationException("Excel-DNA returned no task pane.");

                pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.Width = DefaultWidth;
                pane.VisibleStateChange += OnVisibleStateChange;
                var entry = new PaneEntry(windowHandle, control, pane);
                Entries.Add(windowHandle, entry); // retain both strongly; the Excel-DNA factory uses weak references
                return entry;
            }
            catch (Exception ex)
            {
                if (pane != null)
                {
                    try { pane.VisibleStateChange -= OnVisibleStateChange; } catch { }
                    try { pane.Delete(); } catch { }
                }
                control.Detach();
                control.Dispose();
                Dialogs.Info("The docked Power Tools Hub could not be created:\n" + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Opens the discoverable Comparison workspace and its validated Review picker in the active Excel window.
        /// The Review is loaded into the task pane first; opening a protected Excel copy remains an explicit action.
        /// </summary>
        public static void OpenComparisonReviewPicker()
        {
            SetVisible(true);
            if (!TryGetActiveWindow(out int windowHandle, out _) ||
                !Entries.TryGetValue(windowHandle, out PaneEntry entry))
                return;
            try { entry.Control.OpenComparisonReviewPicker(); }
            catch (Exception ex)
            {
                if (IsDisconnected(ex)) RetireManagedEntry(windowHandle, entry);
                else Dialogs.Info("The Comparison workspace could not be opened:\n" + ex.Message);
            }
        }

        private static void OnVisibleStateChange(CustomTaskPane pane)
        {
            PaneEntry matched = null;
            foreach (PaneEntry entry in Entries.Values)
            {
                if (!ReferenceEquals(entry.Pane, pane)) continue;
                matched = entry;
                break;
            }

            if (matched != null)
            {
                try { if (pane.Visible) matched.Control.RefreshState(); }
                catch (Exception ex)
                {
                    if (IsDisconnected(ex)) RetireManagedEntry(matched.WindowHandle, matched);
                }
            }
            RaiseVisibilityChanged();
        }

        private static bool TryGetActiveWindow(out int windowHandle, out object window)
        {
            windowHandle = 0;
            window = null;
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic activeWindow = app?.ActiveWindow;
                if (activeWindow == null) return false;
                windowHandle = Convert.ToInt32(activeWindow.Hwnd);
                window = activeWindow;
                return windowHandle != 0;
            }
            catch { return false; }
        }

        private static void RaiseVisibilityChanged()
        {
            try { VisibilityChanged?.Invoke(); } catch { }
        }

        private static void TryCreateAndShow(int windowHandle, object window)
        {
            PaneEntry replacement = CreatePane(windowHandle, window);
            if (replacement == null) return;
            try
            {
                replacement.Control.RefreshState();
                replacement.Pane.Visible = true;
            }
            catch (Exception ex)
            {
                if (IsDisconnected(ex)) RetireManagedEntry(windowHandle, replacement);
                else Dialogs.Info("Excel is temporarily busy; the Power Tools Hub could not be opened:\n" + ex.Message);
            }
        }

        private static void SweepClosedWindows()
        {
            if (Entries.Count == 0) return;
            var liveHandles = new HashSet<int>();
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic windows = app.Windows;
                int count = (int)windows.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic window = windows(i);
                    liveHandles.Add(Convert.ToInt32(window.Hwnd));
                }
            }
            catch { return; } // a transient COM failure must never prune a live pane

            var stale = new List<int>();
            foreach (int handle in Entries.Keys)
                if (!liveHandles.Contains(handle)) stale.Add(handle);
            foreach (int handle in stale)
                if (Entries.TryGetValue(handle, out PaneEntry entry)) RetireManagedEntry(handle, entry);
        }

        private static bool IsDisconnected(Exception ex)
        {
            if (ex is InvalidComObjectException) return true;
            if (!(ex is COMException com)) return false;
            int code = com.ErrorCode;
            return code == unchecked((int)0x80010108) || // RPC_E_DISCONNECTED
                   code == unchecked((int)0x800401FD) || // CO_E_OBJNOTCONNECTED
                   code == unchecked((int)0x80010007);   // RPC_E_SERVER_DIED
        }

        private static void RetireManagedEntry(int handle, PaneEntry entry)
        {
            if (Entries.TryGetValue(handle, out PaneEntry current) && ReferenceEquals(current, entry))
                Entries.Remove(handle);
            try { entry.Control.Detach(); } catch { }
            try { entry.Control.Dispose(); } catch { }
        }

        private sealed class PaneEntry
        {
            public PaneEntry(int windowHandle, PowerToolsPane control, CustomTaskPane pane)
            {
                WindowHandle = windowHandle;
                Control = control;
                Pane = pane;
            }

            public int WindowHandle { get; }
            public PowerToolsPane Control { get; }
            public CustomTaskPane Pane { get; }
        }
    }
}
