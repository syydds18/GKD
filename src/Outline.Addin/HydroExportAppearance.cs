using System;
namespace Outline.Addin
{
    internal sealed class HydroPrintCommunicationScope : IDisposable
    {
        private readonly dynamic app;
        private readonly bool originallyEnabled;
        internal HydroPrintCommunicationScope(dynamic app)
        {
            this.app = app;
            originallyEnabled = Convert.ToBoolean(app.PrintCommunication);
            if (originallyEnabled) app.PrintCommunication = false;
        }
        public void Dispose()
        {
            try { app.PrintCommunication = true; if (!originallyEnabled) app.PrintCommunication = false; }
            catch { }
        }
    }
    internal static class HydroExportAppearance
    {
        internal static void ClearFills(dynamic sheet)
        {
            // Conditional formatting can reapply coloured fills even after Interior is cleared.
            // The exported workbook is a static report; source formatting remains untouched.
            sheet.Cells.FormatConditions.Delete();
            sheet.Cells.Interior.Pattern = -4142; // xlNone, not white fill
            sheet.Cells.Interior.ColorIndex = -4142;
        }
    }
}
