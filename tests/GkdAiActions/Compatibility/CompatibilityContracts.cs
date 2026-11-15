#if COMPATIBILITY
using System;
namespace Outline.Addin
{
    internal static class GkdCustomFormulaFunctions
    {
        internal static bool IsCustomName(string name) { return false; }
        internal static bool IsAiFormula(string name) { return false; }
    }
    // Build-only external contracts; the real scheduler/UDF are owned by the integration project.
    internal sealed class GkdAiScheduler { internal void CancelAll() {} }
    internal static class GkdAiFunctions
    {
        internal static readonly GkdAiScheduler Scheduler = new GkdAiScheduler();
        internal static void RefreshCell(string identity) {}
        internal static string CellIdentity(IntPtr sheet, int row, int column) { return ""; }
    }
}
#endif
