using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Outline.Setup
{
    internal static class Installer
    {
        internal const string ProductKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\ExcelOutlineAddin";
        internal static string InstallRoot { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelOutlineAddin", "1.0"); } }
        internal static string StatePath { get { return Path.Combine(InstallRoot, "install-state.xml"); } }
        internal static string PayloadRoot { get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "payload"); } }
        internal static readonly string[] RequiredFiles = { "Outline.xll", "Outline.dna", "Outline.Addin.dll", "Outline.Core.dll", "ExcelDna.Integration.dll", "Newtonsoft.Json.dll" };
        private static string UninstallCommand { get { return "\"" + Path.Combine(InstallRoot, "Outline.Setup.exe") + "\" /uninstall"; } }
        private static RegistryView UninstallView { get { return Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32; } }

        internal static void EnsureExcelClosed()
        {
            int running;
            try { running = ExcelShutdown.Count(); }
            catch (Exception ex) { throw new InvalidOperationException("无法检查当前用户、当前会话的 Excel 是否正在运行：" + ex.Message, ex); }
            if (running != 0) throw new InvalidOperationException("检测到当前用户、当前会话的 Excel 正在运行。请保存并关闭后重试；仅卸载正常关闭超时后，才可另行明确选择强制关闭。不会自动强制结束 Excel。");
        }
        internal static void Install(ExcelTarget target, Action<string> status)
        {
            EnsureExcelClosed();
            // Re-read the PE before registering; never infer Excel bitness from Windows.
            var current = ExcelDiscovery.Inspect(target.ExePath);
            if (current.Architecture != target.Architecture || current.OfficeVersion != target.OfficeVersion)
                throw new InvalidOperationException("Excel 安装信息已变化，请重新检测。");
            if (File.Exists(StatePath)) throw new InvalidOperationException("已存在安装或未完成的安装。请先卸载，再重新安装（可重新选择 Excel 版本）。");
            if (Directory.Exists(InstallRoot) && Directory.EnumerateFileSystemEntries(InstallRoot).Any())
                throw new InvalidOperationException("目标目录包含未知文件且缺少安装状态。为避免覆盖，请先备份并移走该目录。");
            var sources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string architecture in new[] { "x86", "x64" })
            {
                foreach (var file in RequiredFiles)
                {
                    string relative = architecture + "/" + file;
                    string source = Path.Combine(PayloadRoot, architecture, file);
                    if (!File.Exists(source) || new FileInfo(source).Length == 0) throw new InvalidDataException("安装包缺少或为空：payload/" + relative);
                    sources.Add(relative, source);
                }
                using (var stream = File.OpenRead(sources[architecture + "/Outline.xll"]))
                    if (CoreUtil.ReadPeArchitecture(stream) != architecture) throw new InvalidDataException(architecture + " 目录的 XLL 位数不匹配。");
            }
            string workflowKey = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "gkd-workflow.key");
            if (File.Exists(workflowKey)) sources.Add("gkd-workflow.key", workflowKey);
            sources.Add("Outline.Setup.exe", Application.ExecutablePath);
            if (File.Exists(Application.ExecutablePath + ".config")) sources.Add("Outline.Setup.exe.config", Application.ExecutablePath + ".config");
            RejectReparse(InstallRoot);
            Directory.CreateDirectory(InstallRoot);
            RejectReparse(InstallRoot);
            var state = new InstallState { Architecture = target.Architecture, OfficeVersion = target.OfficeVersion, Complete = false, UninstallCommand = UninstallCommand };
            StateFile.Save(StatePath, state);
            try
            {
                foreach (var pair in sources)
                {
                    string destination = Path.Combine(InstallRoot, pair.Key.Replace('/', Path.DirectorySeparatorChar));
                    state.Files.Add(pair.Key); StateFile.Save(StatePath, state); // write-ahead file ownership
                    Directory.CreateDirectory(Path.GetDirectoryName(destination)); RejectReparse(Path.GetDirectoryName(destination));
                    File.Copy(pair.Value, destination, false);
                    if (Hash(pair.Value) != Hash(destination)) throw new IOException("复制后校验失败：" + pair.Key);
                }
                EnsureExcelClosed();
                string expected = CoreUtil.OpenValue(Path.Combine(InstallRoot, target.Architecture, "Outline.xll"));
                using (var options = new RegistryOptions(target.OfficeVersion, target.View, true))
                {
                    var entry = RegistrationLogic.Add(options, target.OfficeVersion, target.View.ToString(), expected,
                        r => { state.Registrations.Add(r); StateFile.Save(StatePath, state); });
                    status("已登记 Excel " + target.OfficeVersion + " / " + target.Architecture + " / " + entry.Name);
                }
                WriteUninstallEntry(state);
                state.Complete = true; StateFile.Save(StatePath, state);
                status("安装完成。请自行打开 Excel；安装程序不会自动启动 Excel。");
            }
            catch
            {
                // Retain the exact journal and copied uninstaller for safe recovery; never lose ownership on a partial failure.
                status("安装未完成，已保留恢复记录。请点击卸载清理后重试。");
                throw;
            }
        }
        internal static void Uninstall(Action<string> status)
        {
            EnsureExcelClosed();
            if (!File.Exists(StatePath)) { status("未找到安装记录；没有修改任何注册项或文件。"); return; }
            RejectReparse(InstallRoot);
            var state = StateFile.Load(StatePath);
            ValidateState(state);
            foreach (var entry in state.Registrations)
            {
                using (var options = new RegistryOptions(entry.OfficeVersion, ParseView(entry.View), false))
                {
                    bool removed = RegistrationLogic.Remove(options, entry);
                    status(removed ? "已移除自己的 " + entry.Name : "保留不存在或已被修改的 " + entry.Name);
                }
            }
            foreach (string relative in state.Files)
            {
                string path = Path.Combine(InstallRoot, relative.Replace('/', Path.DirectorySeparatorChar));
                RejectReparse(Path.GetDirectoryName(path));
                if (File.Exists(path)) { RejectReparse(path); File.Delete(path); }
                if (File.Exists(path)) throw new IOException("无法删除：" + relative);
            }
            RemoveUninstallEntry(state);
            File.Delete(StatePath);
            if (File.Exists(StatePath + ".new")) File.Delete(StatePath + ".new");
            foreach (var architecture in new[] { "x86", "x64" })
            {
                string path = Path.Combine(InstallRoot, architecture);
                if (Directory.Exists(path) && !Directory.EnumerateFileSystemEntries(path).Any()) Directory.Delete(path);
            }
            if (!Directory.EnumerateFileSystemEntries(InstallRoot).Any()) Directory.Delete(InstallRoot);
            status(Directory.Exists(InstallRoot) ? "卸载完成。目录中的未知文件已保留。" : "卸载完成，已清理自己的文件和注册项。");
        }
        private static void ValidateState(InstallState state)
        {
            if (state == null || state.ProductId != "ExcelOutlineAddin/1.0" || state.UninstallCommand != UninstallCommand ||
                (state.Architecture != "x86" && state.Architecture != "x64") || !ExcelDiscovery.Versions.Contains(state.OfficeVersion) || state.Files == null || state.Registrations == null)
                throw new InvalidDataException("安装状态无效；为避免误删，已停止卸载。");
            var allowed = new HashSet<string>(StringComparer.Ordinal);
            foreach (string a in new[] { "x86", "x64" }) foreach (string f in RequiredFiles) allowed.Add(a + "/" + f);
            allowed.Add("Outline.Setup.exe"); allowed.Add("Outline.Setup.exe.config"); allowed.Add("gkd-test-workflow.key"); allowed.Add("gkd-workflow.key");
            foreach (string file in state.Files)
                if (!allowed.Contains(CoreUtil.CheckedRelativePath(file))) throw new InvalidDataException("状态文件包含未知文件，已停止卸载。");
            foreach (var r in state.Registrations)
            {
                if (r.OfficeVersion != state.OfficeVersion || r.View != (state.Architecture == "x64" ? "Registry64" : "Registry32") ||
                    r.Expected != CoreUtil.OpenValue(Path.Combine(InstallRoot, state.Architecture, "Outline.xll")) ||
                    !System.Text.RegularExpressions.Regex.IsMatch(r.Name ?? "", @"\AOPEN(?:[1-9][0-9]*)?\z"))
                    throw new InvalidDataException("状态文件的注册归属无效，已停止卸载。");
            }
        }
        private static RegistryView ParseView(string value) { return value == "Registry64" ? RegistryView.Registry64 : RegistryView.Registry32; }
        private static string Hash(string path) { using (var sha = SHA256.Create()) using (var stream = File.OpenRead(path)) return Convert.ToBase64String(sha.ComputeHash(stream)); }
        private static void RejectReparse(string path)
        {
            // Refuse directory junctions/symlinks anywhere in the destination chain.
            var current = new DirectoryInfo(path);
            while (current != null)
            {
                if ((File.Exists(current.FullName) || Directory.Exists(current.FullName)) && (File.GetAttributes(current.FullName) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("目标路径包含链接或目录联接，已停止操作：" + current.FullName);
                current = current.Parent;
            }
        }
        private static void WriteUninstallEntry(InstallState state)
        {
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, UninstallView))
            {
                using (var old = root.OpenSubKey(ProductKey))
                    if (old != null && !CoreUtil.OwnsValue(state.UninstallCommand, old.GetValue("UninstallString")))
                        throw new InvalidOperationException("同名卸载登记已被其他程序占用，已停止覆盖。");
                using (var key = root.CreateSubKey(ProductKey))
                {
                    key.SetValue("DisplayName", "搞快点", RegistryValueKind.String);
                    key.SetValue("DisplayVersion", "1.6.0", RegistryValueKind.String);
                    key.SetValue("Publisher", "不会开花的草", RegistryValueKind.String);
                    key.SetValue("InstallLocation", InstallRoot, RegistryValueKind.String);
                    key.SetValue("DisplayIcon", Path.Combine(InstallRoot, "Outline.Setup.exe"), RegistryValueKind.String);
                    key.SetValue("UninstallString", state.UninstallCommand, RegistryValueKind.String);
                    key.SetValue("NoModify", 1, RegistryValueKind.DWord); key.SetValue("NoRepair", 1, RegistryValueKind.DWord); key.Flush();
                }
                using (var check = root.OpenSubKey(ProductKey))
                    if (check == null || !CoreUtil.OwnsValue(state.UninstallCommand, check.GetValue("UninstallString"))) throw new IOException("卸载登记核对失败。");
            }
        }
        private static void RemoveUninstallEntry(InstallState state)
        {
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, UninstallView))
            {
                bool own;
                using (var key = root.OpenSubKey(ProductKey)) own = key != null && CoreUtil.OwnsValue(state.UninstallCommand, key.GetValue("UninstallString"));
                if (!own) return;
                root.DeleteSubKeyTree(ProductKey, false);
                using (var check = root.OpenSubKey(ProductKey)) if (check != null) throw new IOException("卸载登记删除后核对失败。");
            }
        }
    }
}
