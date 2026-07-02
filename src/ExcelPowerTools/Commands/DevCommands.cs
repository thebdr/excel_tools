using System;
using System.Diagnostics;
using System.IO;
using ExcelDna.Integration;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// "Framework" ribbon commands — drive the dev loop from inside Excel (the Alt+F11 equivalent):
    /// jump into Visual Studio, open the repo / build-output folders, and show diagnostics.
    /// </summary>
    public static class DevCommands
    {
        [ExcelCommand(Name = "PT_OpenInVisualStudio")]
        public static void OpenSolutionInVisualStudio()
        {
            string sln = FrameworkPaths.SolutionFile;
            if (string.IsNullOrEmpty(sln) || !File.Exists(sln))
            {
                Dialogs.Info("Couldn't locate the solution file.\nRepo root: " +
                             (FrameworkPaths.RepoRoot ?? "(unknown)"));
                return;
            }

            // Prefer the newest installed Visual Studio (via vswhere) rather than whatever version
            // owns the .sln file association, which may be an older VS.
            string devenv = VisualStudio.LatestDevenv();
            if (!string.IsNullOrEmpty(devenv))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = devenv,
                        Arguments = "\"" + sln + "\"",
                        UseShellExecute = false,
                    });
                    return;
                }
                catch { /* fall through to the shell default below */ }
            }

            OpenWithShell(sln);   // fallback: whatever VS owns the .sln association
        }

        [ExcelCommand(Name = "PT_OpenRepoFolder")]
        public static void OpenRepoFolder()
        {
            string root = FrameworkPaths.RepoRoot;
            if (string.IsNullOrEmpty(root) || !Directory.Exists(root))
            {
                Dialogs.Info("Repo root not found.");
                return;
            }
            OpenWithShell(root);
        }

        [ExcelCommand(Name = "PT_OpenBuildOutput")]
        public static void OpenBuildOutputFolder()
        {
            string dir = FrameworkPaths.BuildOutputDir;
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
            {
                Dialogs.Info("Build-output folder not found.");
                return;
            }
            OpenWithShell(dir);
        }

        [ExcelCommand(Name = "PT_About")]
        public static void About()
        {
            var name = typeof(DevCommands).Assembly.GetName();
            Dialogs.Info(
                "Excel Power Tools — framework\n\n" +
                $"Version:    {name.Version}\n" +
                $"Loaded xll: {FrameworkPaths.XllPath}\n" +
                $"Repo root:  {FrameworkPaths.RepoRoot ?? "(unknown)"}\n" +
                $"Solution:   {FrameworkPaths.SolutionFile ?? "(not found)"}\n" +
                $"Runtime:    .NET {Environment.Version}");
        }

        private static void OpenWithShell(string path)
        {
            try
            {
                // UseShellExecute: a .sln opens in Visual Studio, a folder opens in Explorer.
                Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Dialogs.Info("Couldn't open:\n" + path + "\n\n" + ex.Message);
            }
        }
    }
}
