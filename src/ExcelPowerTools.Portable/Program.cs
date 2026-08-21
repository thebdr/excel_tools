using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelPowerTools.Portable
{
    /// <summary>
    /// The portable session launcher: starts a fresh Excel instance and loads the packed
    /// ExcelPowerTools add-in into it with Application.RegisterXLL — for that session only.
    /// Nothing is installed, nothing touches the registry, no admin rights: the user
    /// double-clicks this exe next to the packed .xll and gets the full add-in (ribbon, pane,
    /// commands, events) until they close that Excel. Their normally-installed add-ins and any
    /// already-open Excel windows are untouched — this is a dedicated instance.
    /// </summary>
    internal static class Program
    {
        private const string Title = "Excel PowerTools — portable session";

        [STAThread]
        private static int Main()
        {
            Application.EnableVisualStyles();
            try
            {
                return Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not start the portable session:\n\n" + ex.Message,
                    Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        private static int Run()
        {
            dynamic excel = CreateExcel();
            if (excel == null) return 1;

            try
            {
                bool excel64 = IsExcel64Bit(excel);
                string xllPath = FindAddIn(excel64);
                if (xllPath == null)
                {
                    TryQuit(excel);
                    return 1;
                }
                if (!OfferToClearDownloadBlock(xllPath))
                {
                    TryQuit(excel);
                    return 1;
                }

                excel.Visible = true;
                if ((int)excel.Workbooks.Count == 0) excel.Workbooks.Add();

                while (!(bool)excel.RegisterXLL(xllPath))
                {
                    DialogResult retry = MessageBox.Show(
                        "Excel refused to load the add-in:\n\n" + xllPath +
                        "\n\nThis is the Trust Center gate. Inside the Excel window that just " +
                        "opened, check:\n" +
                        "  •  File → Options → Trust Center → Trust Center Settings → Add-ins —\n" +
                        "      \"Require Application Add-ins to be signed\" must be off\n" +
                        "  •  Protected View / \"block unsigned add-ins\" policies set by IT\n" +
                        "  •  The file still carries a downloaded-from-internet mark\n" +
                        "      (right-click the .xll → Properties → Unblock)\n\n" +
                        "Retry after adjusting?",
                        Title, MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
                    if (retry != DialogResult.Retry)
                    {
                        // The Excel window stays — the user may already be working in it.
                        return 1;
                    }
                }

                // Hand the instance over to the user: with UserControl set, Excel keeps running
                // after this process releases its COM references and exits.
                excel.UserControl = true;
                return 0;
            }
            finally
            {
                if (excel is object) Marshal.FinalReleaseComObject(excel);
            }
        }

        private static dynamic CreateExcel()
        {
            Type type = Type.GetTypeFromProgID("Excel.Application");
            if (type == null)
            {
                MessageBox.Show("Excel is not installed on this machine — the portable session " +
                    "uses your Excel as its interface.", Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            try
            {
                return Activator.CreateInstance(type);
            }
            catch (COMException ex)
            {
                MessageBox.Show("Excel could not be started:\n\n" + ex.Message,
                    Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// The packed add-in matching the running Excel's bitness, searched next to this exe
        /// (packed first, plain second — the latter works when this exe sits in a dev bin folder
        /// where the unpacked .xll has its dependencies alongside). A wrong-bitness-only folder
        /// gets a precise message instead of a RegisterXLL failure that looks like a trust block.
        /// </summary>
        private static string FindAddIn(bool excel64)
        {
            string folder = AppDomain.CurrentDomain.BaseDirectory;
            string[] wanted = excel64
                ? new[] { "ExcelPowerTools-AddIn64-packed.xll", "ExcelPowerTools-AddIn64.xll" }
                : new[] { "ExcelPowerTools-AddIn-packed.xll", "ExcelPowerTools-AddIn.xll" };
            string[] other = excel64
                ? new[] { "ExcelPowerTools-AddIn-packed.xll", "ExcelPowerTools-AddIn.xll" }
                : new[] { "ExcelPowerTools-AddIn64-packed.xll", "ExcelPowerTools-AddIn64.xll" };

            foreach (string name in wanted)
            {
                string path = Path.Combine(folder, name);
                if (File.Exists(path)) return path;
            }

            foreach (string name in other)
            {
                if (!File.Exists(Path.Combine(folder, name))) continue;
                MessageBox.Show("Your Excel is " + (excel64 ? "64" : "32") + "-bit, but this " +
                    "folder only holds the " + (excel64 ? "32" : "64") + "-bit add-in (" + name +
                    ").\n\nCopy " + wanted[0] + " next to this exe.",
                    Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            MessageBox.Show("No add-in found next to this exe.\n\nExpected " + wanted[0] +
                " in:\n" + folder, Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return null;
        }

        /// <summary>
        /// A file copied from a download carries the mark-of-the-web and Excel refuses it
        /// outright. Clearing it is exactly the right-click → Properties → Unblock action,
        /// offered here with consent so the user is not sent hunting through dialogs.
        /// </summary>
        private static bool OfferToClearDownloadBlock(string xllPath)
        {
            string zone = xllPath + ":Zone.Identifier";
            if (NativeMethods.GetFileAttributes(zone) == uint.MaxValue) return true;   // no mark

            DialogResult answer = MessageBox.Show(
                "The add-in file carries the downloaded-from-internet mark, which makes Excel " +
                "refuse to load it:\n\n" + xllPath + "\n\nUnblock it now (the same as right-click " +
                "→ Properties → Unblock)?",
                Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (answer != DialogResult.Yes) return false;

            if (!NativeMethods.DeleteFile(zone))
            {
                MessageBox.Show("The mark could not be removed (the file may be read-only or on " +
                    "a share that forbids it). Unblock it manually via right-click → Properties.",
                    Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        /// <summary>Excel's bitness, read from its own process (COM exposes no direct flag).</summary>
        private static bool IsExcel64Bit(dynamic excel)
        {
            if (!Environment.Is64BitOperatingSystem) return false;
            try
            {
                var hwnd = new IntPtr((int)excel.Hwnd);
                NativeMethods.GetWindowThreadProcessId(hwnd, out uint pid);
                IntPtr process = NativeMethods.OpenProcess(NativeMethods.QueryLimitedInformation, false, pid);
                if (process == IntPtr.Zero) return true;   // modern default
                try
                {
                    return NativeMethods.IsWow64Process(process, out bool wow64) && !wow64;
                }
                finally
                {
                    NativeMethods.CloseHandle(process);
                }
            }
            catch
            {
                return true;   // modern default
            }
        }

        private static void TryQuit(dynamic excel)
        {
            try
            {
                excel.DisplayAlerts = false;
                excel.Quit();
            }
            catch { }
        }

        private static class NativeMethods
        {
            public const uint QueryLimitedInformation = 0x1000;

            [DllImport("user32.dll")]
            public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

            [DllImport("kernel32.dll")]
            public static extern IntPtr OpenProcess(uint access, bool inherit, uint processId);

            [DllImport("kernel32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool IsWow64Process(IntPtr process, [MarshalAs(UnmanagedType.Bool)] out bool wow64);

            [DllImport("kernel32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool CloseHandle(IntPtr handle);

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool DeleteFile(string fileName);

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern uint GetFileAttributes(string fileName);
        }
    }
}
