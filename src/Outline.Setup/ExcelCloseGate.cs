using System;
using System.Threading.Tasks;

namespace Outline.Setup
{
    internal enum ExcelCloseResult { Closed, Cancelled, StillRunning }

    internal static class ExcelCloseGate
    {
        internal static async Task<bool> EnsureClosed(Func<int> count, Func<int, bool> confirm, Action requestClose, Func<Task> wait, int attempts = 120)
        {
            return await CloseNormally(count, confirm, requestClose, wait, attempts) == ExcelCloseResult.Closed;
        }

        internal static async Task<ExcelCloseResult> CloseNormally(Func<int> count, Func<int, bool> confirm, Action requestClose, Func<Task> wait, int attempts = 120)
        {
            if (attempts < 0) throw new ArgumentOutOfRangeException(nameof(attempts));
            int running = count();
            if (running == 0) return ExcelCloseResult.Closed;
            if (!confirm(running)) return ExcelCloseResult.Cancelled;
            requestClose();
            return await WaitUntilClosed(count, wait, attempts) ? ExcelCloseResult.Closed : ExcelCloseResult.StillRunning;
        }

        // Only called by the explicit force button. Never escalate automatically on timeout.
        internal static async Task<bool> ForceClosed(Func<int> count, Func<int, bool> confirmLoss, Action forceClose, Func<Task> wait, int attempts = 20)
        {
            if (attempts < 0) throw new ArgumentOutOfRangeException(nameof(attempts));
            int running = count();
            if (running == 0) return true;
            if (!confirmLoss(running)) return false;
            forceClose();
            return await WaitUntilClosed(count, wait, attempts);
        }

        private static async Task<bool> WaitUntilClosed(Func<int> count, Func<Task> wait, int attempts)
        {
            for (int i = 0; i < attempts; i++) { if (count() == 0) return true; await wait(); }
            return count() == 0;
        }
    }
}
