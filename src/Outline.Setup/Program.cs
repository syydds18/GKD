using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Outline.Setup
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                // Windows cannot delete a running EXE: run the installed maintenance UI from a unique temporary copy.
                if (string.Equals(Path.GetFullPath(Application.ExecutablePath), Path.Combine(Installer.InstallRoot, "Outline.Setup.exe"), StringComparison.OrdinalIgnoreCase))
                {
                    string temp = Path.Combine(Path.GetTempPath(), "ExcelOutline-Uninstall-" + Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(temp);
                    string helper = Path.Combine(temp, "Outline.Setup.exe");
                    File.Copy(Application.ExecutablePath, helper);
                    if (File.Exists(Application.ExecutablePath + ".config")) File.Copy(Application.ExecutablePath + ".config", helper + ".config");
                    var process = Process.Start(new ProcessStartInfo(helper, "/uninstall /wait:" + Process.GetCurrentProcess().Id) { UseShellExecute = false, WorkingDirectory = temp });
                    if (process == null) throw new IOException("无法启动卸载维护程序。");
                    process.Dispose();
                    return 0;
                }
                foreach (string arg in args)
                    if (arg.StartsWith("/wait:", StringComparison.OrdinalIgnoreCase))
                    {
                        int pid;
                        if (int.TryParse(arg.Substring(6), out pid))
                        {
                            try { using (var old = Process.GetProcessById(pid)) if (!old.WaitForExit(30000)) throw new IOException("原安装程序尚未退出，请稍后重试。"); }
                            catch (ArgumentException) { }
                        }
                    }
                using (var mutex = new Mutex(false, @"Local\ExcelOutlineAddin.Setup.1.0"))
                {
                    bool locked;
                    try { locked = mutex.WaitOne(0); } catch (AbandonedMutexException) { locked = true; }
                    if (!locked) { MessageBox.Show("另一个安装或卸载程序正在运行。请先关闭它。", "搞快点", MessageBoxButtons.OK, MessageBoxIcon.Information); return 1; }
                    try { Application.Run(new MainForm(args.Any(a => string.Equals(a, "/uninstall", StringComparison.OrdinalIgnoreCase)))); }
                    finally { mutex.ReleaseMutex(); }
                }
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "无法继续", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }
    }
}
