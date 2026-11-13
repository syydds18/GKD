using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace Outline.Setup
{
    public sealed class ExcelTarget
    {
        public string ExePath { get; set; }
        public string Architecture { get; set; }
        public string OfficeVersion { get; set; }
        public RegistryView View { get { return Architecture == "x64" ? RegistryView.Registry64 : RegistryView.Registry32; } }
        public override string ToString() { return "Excel " + OfficeVersion + " · " + Architecture + "   —   " + ExePath; }
    }

    internal static class ExcelDiscovery
    {
        internal static readonly string[] Versions = { "16.0", "15.0", "14.0" };
        internal static List<ExcelTarget> Find(Action<string> status = null)
        {
            var results = new List<ExcelTarget>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Action<string> add = value => Add(results, seen, ExcelPathCandidates.Executable(value), status);
            Action<string> addRoot = value => { foreach (var candidate in ExcelPathCandidates.InstallRoot(value)) add(candidate); };
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
            {
                try
                {
                    using (var root = RegistryKey.OpenBaseKey(hive, view))
                    {
                        // Each key is isolated: one denied/stale source must not hide all later sources.
                        add(Read(root, @"Software\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE", null, status));
                        addRoot(Read(root, @"Software\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE", "Path", status));
                        foreach (var version in Versions)
                        {
                            addRoot(Read(root, @"Software\Microsoft\Office\" + version + @"\Excel\InstallRoot", "Path", status));
                            addRoot(Read(root, @"Software\Microsoft\Office\" + version + @"\Common\InstallRoot", "Path", status));
                            addRoot(Read(root, @"Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\" + version + @"\Excel\InstallRoot", "Path", status));
                        }
                        foreach (var key in new[] { @"Software\Microsoft\Office\ClickToRun\Configuration", @"Software\Microsoft\Office\15.0\ClickToRun\Configuration" })
                        foreach (var name in new[] { "InstallationPath", "InstallPath" }) addRoot(Read(root, key, name, status));
                        foreach (var progId in new[] { "Excel.Sheet.12", "Excel.Sheet.8", "Excel.SheetMacroEnabled.12", @"Applications\EXCEL.EXE" })
                            add(Read(root, @"Software\Classes\" + progId + @"\shell\Open\command", null, status));
                        add(Read(root, @"Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32", null, status));
                        string clsid = Read(root, @"Software\Classes\Excel.Application\CLSID", null, status);
                        Guid parsed;
                        if (Guid.TryParse(clsid, out parsed))
                            add(Read(root, @"Software\Classes\CLSID\" + parsed.ToString("B") + @"\LocalServer32", null, status));
                        const string uninstall = @"Software\Microsoft\Windows\CurrentVersion\Uninstall";
                        foreach (var subkey in SubKeys(root, uninstall, status))
                        {
                            string key = uninstall + "\\" + subkey;
                            string name = Read(root, key, "DisplayName", status) ?? "";
                            if (name.IndexOf("Office", StringComparison.OrdinalIgnoreCase) < 0 && name.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) < 0 && name.IndexOf("Microsoft 365", StringComparison.OrdinalIgnoreCase) < 0) continue;
                            addRoot(Read(root, key, "InstallLocation", status));
                            add(Read(root, key, "DisplayIcon", status));
                        }
                    }
                }
                catch (Exception ex) when (IsDiscoveryFailure(ex)) { if (status != null) status("跳过无法读取的注册表视图：" + ex.Message); }
            }
            // Running Excel provides an exact path for relocated/portable installs. Never start it.
            try
            {
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                using (process)
                {
                    try { add(process.MainModule.FileName); }
                    catch (Exception ex) when (IsDiscoveryFailure(ex) || ex is InvalidOperationException) { if (status != null) status("无法读取运行中 Excel 的路径：" + ex.Message); }
                }
            }
            catch (Exception ex) when (IsDiscoveryFailure(ex) || ex is InvalidOperationException) { if (status != null) status("跳过运行中 Excel 路径检测：" + ex.Message); }
            foreach (var name in new[] { "ProgramW6432", "ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA" })
            {
                string folder = Environment.GetEnvironmentVariable(name);
                if (!string.IsNullOrWhiteSpace(folder))
                    foreach (var office in new[] { "Microsoft Office", "Microsoft Office 15", "Microsoft Office 14" })
                        addRoot(folder.TrimEnd('\\') + "\\" + office);
            }
            foreach (var folder in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';')) addRoot(folder);
            return results;
        }
        private static string Read(RegistryKey root, string key, string name, Action<string> status)
        {
            try { using (var entry = root.OpenSubKey(key)) return entry == null ? null : entry.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string; }
            catch (Exception ex) when (IsDiscoveryFailure(ex)) { if (status != null) status("跳过注册表来源 " + key + "：" + ex.Message); return null; }
        }
        private static string[] SubKeys(RegistryKey root, string key, Action<string> status)
        {
            try { using (var entry = root.OpenSubKey(key)) return entry == null ? new string[0] : entry.GetSubKeyNames(); }
            catch (Exception ex) when (IsDiscoveryFailure(ex)) { if (status != null) status("跳过注册表来源 " + key + "：" + ex.Message); return new string[0]; }
        }
        private static bool IsDiscoveryFailure(Exception ex)
        {
            return ex is IOException || ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is ArgumentException || ex is System.ComponentModel.Win32Exception || ex is NotSupportedException;
        }
        private static void Add(List<ExcelTarget> results, HashSet<string> seen, string path, Action<string> status)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            try
            {
                path = Path.GetFullPath(path);
                if (seen.Add(path) && File.Exists(path)) results.Add(Inspect(path));
            }
            catch (Exception ex) when (IsDiscoveryFailure(ex)) { if (status != null) status("Excel 候选路径无法使用：" + path + " — " + ex.Message); }
        }
        internal static ExcelTarget Inspect(string path)
        {
            path = Path.GetFullPath(Environment.ExpandEnvironmentVariables(path.Trim().Trim('"')));
            if (!string.Equals(Path.GetFileName(path), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase) || !File.Exists(path))
                throw new InvalidDataException("请选择真实存在的 EXCEL.EXE。");
            string architecture;
            using (var stream = File.OpenRead(path)) architecture = CoreUtil.ReadPeArchitecture(stream);
            int major = FileVersionInfo.GetVersionInfo(path).FileMajorPart;
            string version = major.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".0";
            if (Array.IndexOf(Versions, version) < 0) throw new InvalidDataException("仅支持 Excel 2010 / 2013 / 2016 及更新的 16.0 系列。");
            return new ExcelTarget { ExePath = path, Architecture = architecture, OfficeVersion = version };
        }
    }

    internal sealed class RegistryOptions : IOptionStore, IDisposable
    {
        private readonly RegistryKey root;
        private readonly RegistryKey key;
        internal RegistryOptions(string version, RegistryView view, bool create)
        {
            root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
            string path = @"Software\Microsoft\Office\" + version + @"\Excel\Options";
            key = create ? root.CreateSubKey(path, true) : root.OpenSubKey(path, true);
        }
        public string[] Names() { return key == null ? new string[0] : key.GetValueNames(); }
        public object Get(string name) { return key == null ? null : key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames); }
        public void Set(string name, string value)
        {
            if (key == null) throw new InvalidOperationException("Excel 注册路径不存在。");
            key.SetValue(name, value, RegistryValueKind.String); key.Flush();
        }
        public void Delete(string name) { if (key != null) { key.DeleteValue(name, false); key.Flush(); } }
        public void Dispose() { if (key != null) key.Dispose(); root.Dispose(); }
    }
}
