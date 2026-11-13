using System;

namespace Outline.Setup
{
    internal static class ExcelProcessScope
    {
        internal static bool Matches(string currentSid, int currentSession, string processSid, int processSession)
        {
            return !string.IsNullOrEmpty(currentSid) && currentSession == processSession &&
                string.Equals(currentSid, processSid, StringComparison.Ordinal);
        }
    }
}
