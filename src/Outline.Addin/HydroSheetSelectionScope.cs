using System;
using System.Collections.Generic;
namespace Outline.Addin
{
    // Range.Copy from grouped tabs can place a multi-sheet range on the clipboard.
    // Activate alone does not ungroup: use Select(Replace:true), then restore tabs.
    internal sealed class HydroSheetSelectionScope : IDisposable
    {
        private readonly object workbook;
        private readonly object window;
        private readonly object active;
        private readonly List<object> selected = new List<object>();
        internal HydroSheetSelectionScope(dynamic app, dynamic workbook)
        {
            this.workbook = (object)workbook;
            dynamic sourceWindow = app.ActiveWindow;
            window = (object)sourceWindow;
            active = (object)app.ActiveSheet;
            foreach (dynamic sheet in sourceWindow.SelectedSheets) selected.Add((object)sheet);
        }
        internal void SelectOnly(dynamic sheet)
        {
            ((dynamic)workbook).Activate();
            sheet.Select(true);
        }
        public void Dispose()
        {
            // A closed sheet/window must not hide the actual export exception.
            try
            {
                ((dynamic)workbook).Activate();
                if (window != null) { try { ((dynamic)window).Activate(); } catch { } }
                bool replace = true;
                foreach (dynamic sheet in selected)
                {
                    sheet.Select(replace);
                    replace = false;
                }
                if (active != null) ((dynamic)active).Activate();
            }
            catch { }
        }
    }
}
