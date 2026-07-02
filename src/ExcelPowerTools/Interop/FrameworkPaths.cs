using System;
using System.IO;
using System.Linq;
using System.Reflection;
using ExcelDna.Integration;

namespace ExcelPowerTools.Interop
{
    /// <summary>
    /// Locates the framework's own source/repo at runtime, so the "Framework" ribbon commands can
    /// open the solution in Visual Studio, the repo folder, etc. — the C# equivalent of Alt+F11.
    /// </summary>
    internal static class FrameworkPaths
    {
        /// <summary>The loaded add-in's .xll path.</summary>
        public static string XllPath => ExcelDnaUtil.XllPath;

        /// <summary>Folder the .xll was loaded from (the build output — bin\... or publish\...).</summary>
        public static string BuildOutputDir
        {
            get
            {
                try { return Path.GetDirectoryName(XllPath); } catch { return null; }
            }
        }

        /// <summary>
        /// The repository root. Primary source is the RepoRoot value baked into the assembly at build
        /// time; falls back to walking up from the loaded .xll to a folder that has a .sln/.slnx, or a
        /// .git entry (a directory in a normal clone, or a file in a worktree/submodule).
        /// </summary>
        public static string RepoRoot
        {
            get
            {
                string baked = typeof(FrameworkPaths).Assembly
                    .GetCustomAttributes<AssemblyMetadataAttribute>()
                    .FirstOrDefault(a => a.Key == "RepoRoot")?.Value;
                if (!string.IsNullOrEmpty(baked) && Directory.Exists(baked)) return baked;

                return FindUpwards(BuildOutputDir, IsRepoRoot);
            }
        }

        /// <summary>The solution at the repo root (.sln preferred, then .slnx).</summary>
        public static string SolutionFile
        {
            get
            {
                string root = RepoRoot;
                if (string.IsNullOrEmpty(root) || !Directory.Exists(root)) return null;
                return FirstFileWithExtension(root, ".sln") ?? FirstFileWithExtension(root, ".slnx");
            }
        }

        private static bool IsRepoRoot(string dir)
        {
            return FirstFileWithExtension(dir, ".sln") != null
                || FirstFileWithExtension(dir, ".slnx") != null
                || Directory.Exists(Path.Combine(dir, ".git"))
                || File.Exists(Path.Combine(dir, ".git"));   // worktree / submodule: .git is a file
        }

        /// <summary>
        /// First file in <paramref name="dir"/> with exactly <paramref name="ext"/>. We match the
        /// extension ourselves rather than passing "*.sln" to EnumerateFiles, because on .NET
        /// Framework a 3-char search extension also matches longer ones (".sln" would return ".slnx").
        /// </summary>
        private static string FirstFileWithExtension(string dir, string ext)
        {
            try
            {
                return Directory.EnumerateFiles(dir).FirstOrDefault(
                    f => string.Equals(Path.GetExtension(f), ext, StringComparison.OrdinalIgnoreCase));
            }
            catch { return null; }
        }

        private static string FindUpwards(string startDir, Func<string, bool> match)
        {
            try
            {
                for (var dir = new DirectoryInfo(startDir); dir != null; dir = dir.Parent)
                {
                    // Skip a directory we can't inspect rather than abandoning the whole climb.
                    try { if (match(dir.FullName)) return dir.FullName; }
                    catch { /* keep going up */ }
                }
            }
            catch { /* startDir was null/invalid */ }
            return null;
        }
    }
}
