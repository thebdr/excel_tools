using System;
using System.Diagnostics;
using System.IO;

namespace ExcelPowerTools.Interop
{
    /// <summary>Locates the newest installed Visual Studio via vswhere.</summary>
    internal static class VisualStudio
    {
        /// <summary>
        /// Full path to the newest installed VS <c>devenv.exe</c> (including preview builds), or null.
        /// Used so "Open in VS" launches the latest VS rather than whatever version owns the .sln
        /// file association.
        /// </summary>
        public static string LatestDevenv()
        {
            try
            {
                string pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                string vswhere = Path.Combine(pf86, "Microsoft Visual Studio", "Installer", "vswhere.exe");
                if (!File.Exists(vswhere)) return null;

                var psi = new ProcessStartInfo
                {
                    FileName = vswhere,
                    Arguments = "-latest -prerelease -property productPath",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                };

                using (var p = Process.Start(psi))
                {
                    string output = (p.StandardOutput.ReadToEnd() ?? string.Empty).Trim();
                    p.WaitForExit(5000);

                    // productPath for an IDE install is devenv.exe (BuildTools is excluded by default).
                    if (!string.IsNullOrEmpty(output)
                        && output.EndsWith("devenv.exe", StringComparison.OrdinalIgnoreCase)
                        && File.Exists(output))
                        return output;
                }
            }
            catch { /* ignore — caller falls back to the shell default */ }

            return null;
        }
    }
}
