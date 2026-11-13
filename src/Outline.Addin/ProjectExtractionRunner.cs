using System;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using Outline.Core;

namespace Outline.Addin
{
    internal static class ProjectExtractionRunner
    {
        internal static void Run()
        {
            dynamic app = ExcelBridge.App;
            if (app.ActiveWorkbook == null) throw new InvalidOperationException("请先打开需要提取的工作簿。");
            dynamic sourceWorkbook = app.ActiveWorkbook;
            string sourcePath = Convert.ToString(sourceWorkbook.FullName, CultureInfo.InvariantCulture);
            string exportWorkbookName = Path.GetFileNameWithoutExtension(sourcePath);
            if (string.IsNullOrEmpty(sourcePath) ||
                (sourcePath.IndexOf("\\", StringComparison.Ordinal) < 0 && sourcePath.IndexOf("/", StringComparison.Ordinal) < 0))
                throw new InvalidOperationException("当前工作簿尚未保存，请先保存后再执行提项目。");

            var choices = HydropowerExtraction.ReadChoices(app, sourceWorkbook, sourcePath);
            using (var ranges = new HydropowerRangeForm(choices, sourcePath))
            {
                DialogResult result;
                do
                {
                    result = ranges.ShowDialog(new ExcelOwner());
                    if (result == DialogResult.Retry && ranges.RangePickRequested != null)
                        HydropowerRangeForm.PickRange(app, ranges.RangePickRequested);
                }
                while (result == DialogResult.Retry);
                if (result != DialogResult.OK) return;
                var selected = ranges.Selected;
                using (var confirm = new HydroConfirmForm(selected))
                    if (confirm.ShowDialog(new ExcelOwner()) != DialogResult.OK) return;
                bool screen = app.ScreenUpdating;
                bool events = app.EnableEvents;
                bool alerts = app.DisplayAlerts;
                int calculation = Convert.ToInt32(app.Calculation, CultureInfo.InvariantCulture);
                bool calculateBeforeSave = app.CalculateBeforeSave;
                string output;
                try
                {
                    app.Calculation = -4135; // xlCalculationManual: do not recalculate source external links while exporting.
                    app.CalculateBeforeSave = false;
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                    app.DisplayAlerts = false;
                    using (var progress = new HydroProgressForm(selected.Count))
                    {
                        progress.Show(new ExcelOwner());
                        progress.Report(0, "准备导出，共 " + selected.Count + " 张工作表…");
                        output = HydropowerExtraction.Run(app, sourceWorkbook, sourcePath, selected, exportWorkbookName, progress);
                        progress.Report(selected.Count, "导出完成");
                        progress.Close();
                    }
                }
                finally
                {
                    try { app.CalculateBeforeSave = calculateBeforeSave; }
                    finally
                    {
                        try { app.Calculation = calculation; }
                        finally
                        {
                            try { app.ScreenUpdating = screen; }
                            finally
                            {
                                try { app.EnableEvents = events; }
                                finally { app.DisplayAlerts = alerts; }
                            }
                        }
                    }
                }
                using (var success = new HydroSuccessForm(output)) success.ShowDialog(new ExcelOwner());
            }
        }
    }
}
