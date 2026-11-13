using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace Outline.Setup
{
    internal static class ExcelShutdown
    {
        private const uint QueryLimitedInformation = 0x1000, Synchronize = 0x100000, Terminate = 0x0001;
        private delegate bool EnumProc(IntPtr hwnd, IntPtr param);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr param);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int length);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern SafeProcessHandle OpenProcess(uint access, bool inherit, int pid);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool QueryFullProcessImageName(SafeProcessHandle process, uint flags, StringBuilder path, ref int size);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool TerminateProcess(SafeProcessHandle process, uint exitCode);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern uint WaitForSingleObject(SafeProcessHandle process, uint milliseconds);
        [DllImport("advapi32.dll", SetLastError = true)] private static extern bool OpenProcessToken(SafeProcessHandle process, uint access, out IntPtr token);
        [DllImport("advapi32.dll", SetLastError = true)] private static extern bool GetTokenInformation(IntPtr token, int infoClass, out int info, int size, out int needed);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);

        internal sealed class Snapshot : IDisposable
        {
            internal readonly Dictionary<int, SafeProcessHandle> Processes = new Dictionary<int, SafeProcessHandle>();
            internal int Count { get { return Processes.Count; } }
            internal void ForceClose(Action<string> status)
            {
                // Operate on pinned handles, not a fresh PID lookup: never hit a reused PID or new Excel.
                foreach (var entry in Processes)
                {
                    uint state = WaitForSingleObject(entry.Value, 0);
                    if (state == 0) continue;
                    if (state != 258) throw new Win32Exception(Marshal.GetLastWin32Error(), "无法检查 Excel 进程状态；已停止。");
                    if (!TerminateProcess(entry.Value, 1))
                    {
                        if (WaitForSingleObject(entry.Value, 0) == 0) continue;
                        throw new Win32Exception(Marshal.GetLastWin32Error(), "无法结束 Excel PID " + entry.Key + "。请手动关闭；不会提权。");
                    }
                    status("已按明确确认强制结束 Excel PID " + entry.Key + "；未保存的更改可能已丢失。");
                }
            }
            public void Dispose() { foreach (var handle in Processes.Values) handle.Dispose(); Processes.Clear(); }
        }

        private static bool IsOwned(SafeProcessHandle handle, string sid, int session)
        {
            IntPtr token;
            if (!OpenProcessToken(handle, 0x0008, out token)) throw new Win32Exception(Marshal.GetLastWin32Error(), "无法核对 Excel 所属用户；已停止。");
            try
            {
                int tokenSession, needed;
                if (!GetTokenInformation(token, 12, out tokenSession, sizeof(int), out needed)) throw new Win32Exception(Marshal.GetLastWin32Error());
                using (var identity = new WindowsIdentity(token))
                    return identity.User != null && ExcelProcessScope.Matches(sid, session, identity.User.Value, tokenSession);
            }
            finally { CloseHandle(token); }
        }

        internal static Snapshot CaptureForForce() { return Capture(true); }
        private static Snapshot Capture(bool forForce)
        {
            var snapshot = new Snapshot();
            try
            {
                using (var identity = WindowsIdentity.GetCurrent())
                using (var self = Process.GetCurrentProcess())
                {
                    if (identity.User == null) throw new InvalidOperationException("无法核对当前用户；已停止。");
                    var processes = Process.GetProcessesByName("EXCEL");
                    try
                    {
                        foreach (var process in processes)
                        {
                            try
                            {
                                if (process.HasExited || process.SessionId != self.SessionId) continue;
                                var handle = OpenProcess(QueryLimitedInformation | Synchronize | (forForce ? Terminate : 0), false, process.Id);
                                if (handle.IsInvalid)
                                {
                                    int error = Marshal.GetLastWin32Error(); handle.Dispose();
                                    if (error == 87) continue; // exited during enumeration
                                    throw new Win32Exception(error, "无法检查 Excel PID " + process.Id + "；请手动关闭后重试。");
                                }
                                try
                                {
                                    if (WaitForSingleObject(handle, 0) == 0 || !IsOwned(handle, identity.User.Value, self.SessionId)) continue;
                                    var image = new StringBuilder(32768); int size = image.Capacity;
                                    if (!QueryFullProcessImageName(handle, 0, image, ref size))
                                    {
                                        if (WaitForSingleObject(handle, 0) == 0) continue;
                                        throw new Win32Exception(Marshal.GetLastWin32Error(), "无法核对 Excel 程序名称；已停止。");
                                    }
                                    if (!string.Equals(System.IO.Path.GetFileName(image.ToString()), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase)) continue;
                                    snapshot.Processes.Add(process.Id, handle); handle = null;
                                }
                                finally { if (handle != null) handle.Dispose(); }
                            }
                            catch (InvalidOperationException) { if (!process.HasExited) throw; }
                        }
                    }
                    finally { foreach (var process in processes) process.Dispose(); }
                }
                return snapshot;
            }
            catch { snapshot.Dispose(); throw; }
        }

        internal static int Count() { using (var snapshot = Capture(false)) return snapshot.Count; }
        internal static void RequestClose(Action<string> status)
        {
            using (var snapshot = Capture(false))
            {
                int sent = 0;
                bool enumerated = EnumWindows((hwnd, param) => {
                    uint pid; GetWindowThreadProcessId(hwnd, out pid);
                    SafeProcessHandle handle;
                    if (!snapshot.Processes.TryGetValue((int)pid, out handle) || WaitForSingleObject(handle, 0) == 0) return true;
                    var name = new StringBuilder(256); GetClassName(hwnd, name, name.Capacity);
                    if (name.ToString() == "XLMAIN")
                    {
                        if (PostMessage(hwnd, 0x0010, IntPtr.Zero, IntPtr.Zero)) sent++;
                        else status("无法发送关闭请求，请手动关闭该 Excel（可能权限不同）。");
                    }
                    return true;
                }, IntPtr.Zero);
                if (!enumerated) throw new InvalidOperationException("无法枚举 Excel 窗口，已停止；没有强制结束进程。");
                status("已向 " + sent + " 个当前用户、当前会话的 Excel 窗口发送正常关闭请求。请处理保存提示；不会自动放弃更改或强杀进程。");
                if (sent == 0) status("未找到可关闭的 Excel 主窗口；正常关闭等待超时后，可自行处理后台实例，或明确选择强制关闭选项。");
            }
        }
    }
}
