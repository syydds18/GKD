using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outline.Core;

namespace Outline.Addin
{
    internal sealed class HydroSheetChoice : INotifyPropertyChanged
    {
        internal object SourceSheet;
        private bool include;
        private string printRange;
        public event PropertyChangedEventHandler PropertyChanged;
        public bool Include { get { return include; } set { if (include == value) return; include = value; Changed("Include"); } }
        public string SheetName { get; set; }
        public bool IsRecommended { get; set; }
        public bool IsVisible { get; set; }
        public string PrintRange { get { return printRange; } set { if (printRange == value) return; printRange = value; Changed("PrintRange"); } }
        private void Changed(string name) { var handler = PropertyChanged; if (handler != null) handler(this, new PropertyChangedEventArgs(name)); }
    }

    internal sealed class HydropowerRangeForm : Form
    {
        private readonly DataGridView grid;
        private readonly BindingList<HydroSheetChoice> rows;
        private readonly List<HydroSheetChoice> workbookRows;
        private readonly string workbookPath;
        private readonly HydroExportPlanStore planStore = HydroExportPlanStore.ForLocalApplicationData();
        internal List<HydroSheetChoice> Selected { get { return workbookRows.Where(x => x.Include && x.IsVisible).ToList(); } }
        internal HydroSheetChoice RangePickRequested { get; private set; }

        internal HydropowerRangeForm(List<HydroSheetChoice> sheets, string workbookPath)
        {
            workbookRows = sheets.ToList();
            rows = new BindingList<HydroSheetChoice>(sheets.OrderByDescending(x => x.Include).ToList());
            this.workbookPath = workbookPath;
            Ui.Base(this, new System.Drawing.Size(900, 620), "水电工程 — 选择工作表与打印范围");
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 160));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            var planActions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            var savePlan = Ui.Button("保存为下次默认", true);
            savePlan.Click += (s, e) => SaveDefaultPlan();
            var loadPlan = Ui.Button("加载上次导表范围");
            loadPlan.Click += (s, e) => LoadDefaultPlan();
            planActions.Controls.Add(savePlan); planActions.Controls.Add(loadPlan);
            header.Controls.Add(planActions, 0, 0);
            header.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = "使用提项目功能前，建议提前设置分页预览范围，并选择(可多选)需要导出的工作表范围。",
                AutoSize = false,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                ForeColor = System.Drawing.Color.FromArgb(96, 107, 118)
            }, 0, 1);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            var all = Ui.Button("全选");
            all.Click += (s, e) => SetAll(true);
            var none = Ui.Button("取消全选");
            none.Click += (s, e) => SetAll(false);
            var recommended = Ui.Button("推荐报表");
            recommended.Click += (s, e) => ApplyRecommendations();
            actions.Controls.Add(all); actions.Controls.Add(none); actions.Controls.Add(recommended);
            header.Controls.Add(actions, 0, 2);
            layout.Controls.Add(header, 0, 0);
            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                DataSource = rows
            };
            grid.Columns.Add(new DataGridViewCheckBoxColumn { DataPropertyName = "Include", HeaderText = "导出", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, FillWeight = 12, MinimumWidth = 55 });
            var sheetColumn = new DataGridViewTextBoxColumn { DataPropertyName = "SheetName", HeaderText = "工作表", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, FillWeight = 34, MinimumWidth = 160 };
            sheetColumn.DefaultCellStyle.BackColor = System.Drawing.Color.White;
            grid.Columns.Add(sheetColumn);
            grid.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "PrintRange", HeaderText = "打印范围", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, FillWeight = 34, MinimumWidth = 160 });
            grid.Columns.Add(new DataGridViewButtonColumn { Name = "pickRange", HeaderText = "范围", Text = "框选范围", UseColumnTextForButtonValue = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, FillWeight = 20, MinimumWidth = 100 });
            grid.CellFormatting += (s, e) =>
            {
                if (e.RowIndex >= 0 && rows[e.RowIndex].IsRecommended)
                {
                    e.CellStyle.BackColor = System.Drawing.Color.Honeydew;
                    e.CellStyle.SelectionBackColor = System.Drawing.Color.PaleGreen;
                }
            };
            grid.CellContentClick += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == grid.Columns["pickRange"].Index)
                {
                    RangePickRequested = rows[e.RowIndex];
                    DialogResult = DialogResult.Retry;
                }
            };
            grid.MinimumSize = new System.Drawing.Size(600, 180);
            layout.Controls.Add(grid, 0, 1);
            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 8, 0, 0) };
            var cancel = Ui.Button("取消"); cancel.DialogResult = DialogResult.Cancel;
            var go = Ui.Button("继续", true);
            go.Click += (s, e) =>
            {
                grid.EndEdit();
                if (!Selected.Any()) { MessageBox.Show("请至少选择一张工作表。", "水电工程"); return; }
                foreach (var choice in Selected)
                    if (!HydroRangeAddress.IsValid(choice.PrintRange)) { MessageBox.Show("“" + choice.SheetName + "”的范围地址无效。请使用 A1:B20 这样的单一连续区域。", "水电工程"); return; }
                DialogResult = DialogResult.OK;
            };
            footer.Controls.Add(cancel); footer.Controls.Add(go); layout.Controls.Add(footer, 0, 2);
            Controls.Add(layout); AcceptButton = go; CancelButton = cancel;
            MinimumSize = new System.Drawing.Size(760, 480);
            StartPosition = FormStartPosition.CenterScreen;
            HydroWindowBounds.Fit(this);
        }

        private void SetAll(bool include)
        {
            foreach (var row in rows) row.Include = include && row.IsVisible;
            grid.Refresh();
        }

        private void ApplyRecommendations()
        {
            foreach (var row in rows) row.Include = row.IsVisible && row.IsRecommended;
            grid.Refresh();
        }

        private void LoadDefaultPlan()
        {
            grid.EndEdit();
            try
            {
                var plan = planStore.Load(workbookPath);
                if (plan == null || plan.Sheets.Count == 0)
                {
                    MessageBox.Show(this, "没有找到可用的上次导表方案（可能尚未保存、文件损坏或版本不兼容）。当前选择和范围未改变。", "水电工程");
                    return;
                }
                var matches = HydroExportPlanPolicy.MatchVisibleRanges(plan,
                    workbookRows.Where(x => x.IsVisible).Select(x => x.SheetName));
                if (matches.Count == 0)
                {
                    MessageBox.Show(this, "上次方案没有匹配到当前可见工作表的有效范围；工作表可能已改名、删除或隐藏，或范围已失效。当前选择和范围未改变。", "水电工程");
                    return;
                }
                var byName = matches.ToDictionary(x => x.Name, StringComparer.OrdinalIgnoreCase);
                foreach (var row in workbookRows)
                {
                    HydroExportPlanSheet saved;
                    bool matched = row.IsVisible && byName.TryGetValue(row.SheetName, out saved);
                    row.Include = matched;
                    if (matched) row.PrintRange = byName[row.SheetName].PrintRange;
                }
                // Reorder once after an explicit action; avoid moving a row while a
                // checkbox or range cell is being edited.
                rows.RaiseListChangedEvents = false;
                try
                {
                    rows.Clear();
                    foreach (var row in workbookRows.OrderByDescending(x => x.Include)) rows.Add(row);
                }
                finally { rows.RaiseListChangedEvents = true; rows.ResetBindings(); }
                int skipped = plan.Sheets.Count - matches.Count;
                MessageBox.Show(this, "已加载 " + matches.Count + " 张工作表的上次导表范围。" +
                    (skipped > 0 ? "已跳过 " + skipped + " 项未匹配、隐藏、重复或范围无效的记录。" : ""), "水电工程");
            }
            catch (Exception ex) { MessageBox.Show(this, "加载上次导表范围失败：" + ex.Message, "水电工程"); }
        }

        private void SaveDefaultPlan()
        {
            grid.EndEdit();
            var selected = Selected;
            if (!selected.Any())
            {
                MessageBox.Show(this, "请至少选择一张工作表后再保存默认方案。", "水电工程");
                return;
            }
            foreach (var choice in selected)
            {
                if (!HydroRangeAddress.IsValid(choice.PrintRange))
                {
                    MessageBox.Show(this, "“" + choice.SheetName + "”的范围地址无效。", "水电工程");
                    return;
                }
            }
            try
            {
                planStore.Save(workbookPath, selected.Select(x => new HydroExportPlanSheet { Name = x.SheetName, PrintRange = x.PrintRange }));
                MessageBox.Show(this, "已保存工作表和打印范围；下次可点击“加载上次导表范围”恢复。", "水电工程");
            }
            catch (Exception ex) { MessageBox.Show(this, "保存默认方案失败：" + ex.Message, "水电工程"); }
        }

        internal static bool PickRange(dynamic app, HydroSheetChoice choice)
        {
            try
            {
                app.Visible = true;
                dynamic sheet = choice.SourceSheet;
                sheet.Activate();
                object picked = app.InputBox("请在当前工作表上拖选一个连续区域，然后点击确定。", "选择打印范围", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (picked is bool && !(bool)picked) return false;
                dynamic range = picked;
                string address = Convert.ToString(range.Address, CultureInfo.InvariantCulture).Replace("$", "");
                if (!HydroRangeAddress.IsValid(address))
                {
                    MessageBox.Show("请选择单一连续单元格区域。", "水电工程");
                    return false;
                }
                choice.PrintRange = address;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("没有取得所选区域：" + ex.Message, "水电工程");
                return false;
            }
        }
    }

    internal static class HydropowerExtraction
    {
        internal static List<HydroSheetChoice> ReadChoices(dynamic app, dynamic workbook, string workbookPath)
        {
            var result = new List<HydroSheetChoice>();
            var sheets = new List<dynamic>();
            var names = new List<string>();
            foreach (dynamic sheet in workbook.Worksheets)
            {
                sheets.Add((object)sheet);
                names.Add(Convert.ToString(sheet.Name, CultureInfo.InvariantCulture));
            }
            var recommended = HydroReportRecommender.Recommend(names);
            var currentSelection = ReadSelectedSheetNames(app, workbook);
            for (int index = 0; index < sheets.Count; index++)
            {
                dynamic sheet = sheets[index];
                string address = HydroLivePrintArea.Read(app, sheet);
                bool visible = Convert.ToInt32(sheet.Visible, CultureInfo.InvariantCulture) == -1;
                string name = names[index];
                bool isRecommended = recommended.Contains(name);
                bool include = HydroExportPlanPolicy.ShouldInclude(name, isRecommended, currentSelection, null);
                result.Add(new HydroSheetChoice { SourceSheet = (object)sheet, Include = visible && include, IsVisible = visible, IsRecommended = isRecommended, SheetName = name, PrintRange = address });
            }
            return result;
        }

        private static List<string> ReadSelectedSheetNames(dynamic app, dynamic workbook)
        {
            try
            {
                dynamic selected = app.ActiveWindow.SelectedSheets;
                var names = new List<string>();
                foreach (dynamic sheet in selected)
                {
                    if (!object.ReferenceEquals((object)sheet.Parent, (object)workbook)) continue;
                    names.Add(Convert.ToString(sheet.Name, CultureInfo.InvariantCulture));
                }
                return names;
            }
            catch
            {
                try
                {
                    dynamic active = app.ActiveSheet;
                    if (object.ReferenceEquals((object)active.Parent, (object)workbook))
                        return new List<string> { Convert.ToString(active.Name, CultureInfo.InvariantCulture) };
                }
                catch { }
                return null;
            }
        }

        internal static string Run(dynamic app, dynamic sourceWorkbook, string sourcePath, List<HydroSheetChoice> selected, string exportWorkbookName, HydroProgressForm progress)
        {
            using (var selectionScope = new HydroSheetSelectionScope((object)app, (object)sourceWorkbook))
            {
            foreach (var choice in selected)
            {
                dynamic source = choice.SourceSheet;
                if (!object.ReferenceEquals((object)source.Parent, (object)sourceWorkbook)) throw new InvalidOperationException("工作表所属工作簿已变化，请重新开始。");
                if ((bool)source.ProtectContents) throw new InvalidOperationException("工作表“" + choice.SheetName + "”受保护；没有创建结果文件。");
                try { dynamic range = source.Range[choice.PrintRange]; int rows = (int)range.Rows.Count, cols = (int)range.Columns.Count, areas = (int)range.Areas.Count; if (areas != 1) throw new InvalidOperationException("请为每张表指定单一连续区域。"); if (rows < 1 || cols < 1 || (long)rows * cols > 2000000L) throw new InvalidOperationException("打印范围过大（超过 200 万个单元格）。"); }
                catch (Exception ex) { throw new InvalidOperationException("工作表“" + choice.SheetName + "”的范围“" + choice.PrintRange + "”无效：" + ex.Message); }
            }

            string output = NewOutputPath(sourcePath, exportWorkbookName);
            dynamic result = app.Workbooks.Add();
            string placeholderName = Convert.ToString(result.Worksheets[1].Name, CultureInfo.InvariantCulture);
            try
            {
                using (var printBatch = new HydroPrintCommunicationScope((object)app))
                {
                for (int index = 0; index < selected.Count; index++)
                {
                    var choice = selected[index];
                    progress.Report(index + 1, "正在处理：" + choice.SheetName);
                    dynamic source = choice.SourceSheet;
                    dynamic sourcePage = source.PageSetup;
                    dynamic sourceRange = source.Range[choice.PrintRange];
                    dynamic target = result.Worksheets.Add(Type.Missing, result.Worksheets[result.Worksheets.Count]);
                    target.Name = SafeSheetName(choice.SheetName, result);
                    dynamic dest = target.Range["A1"].Resize[sourceRange.Rows.Count, sourceRange.Columns.Count];
                    selectionScope.SelectOnly((object)source);
                    HydroRangeCopy.CopyValuesAndFormats((object)sourceRange, (object)dest);
                    try { app.CutCopyMode = false; } catch { }
                    DeleteComments(target, dest);
                    FormatRange(target, dest);
                    HydroExportAppearance.ClearFills((object)target);
                    HydroOutlineTransfer.Copy((object)sourceRange, (object)target);
                    target.PageSetup.PrintArea = dest.Address;
                    target.PageSetup.Zoom = false;
                    target.PageSetup.FitToPagesWide = 1;
                    target.PageSetup.FitToPagesTall = false;
                    target.PageSetup.CenterHorizontally = true;
                    target.PageSetup.Orientation = sourcePage.Orientation;
                    target.PageSetup.LeftMargin = sourcePage.LeftMargin;
                    target.PageSetup.RightMargin = sourcePage.RightMargin;
                    target.PageSetup.TopMargin = sourcePage.TopMargin;
                    target.PageSetup.BottomMargin = sourcePage.BottomMargin;
                    try { target.PageSetup.PrintTitleRows = sourcePage.PrintTitleRows; } catch { }
                    try { target.PageSetup.PrintTitleColumns = sourcePage.PrintTitleColumns; } catch { }
                }
                } // Commit queued printer settings before SaveAs.
                for (int i = result.Worksheets.Count; i >= 1; i--)
                {
                    dynamic sheet = result.Worksheets[i];
                    if (string.Equals(Convert.ToString(sheet.Name, CultureInfo.InvariantCulture), placeholderName, StringComparison.Ordinal)) sheet.Delete();
                }
            try { result.RemovePersonalInformation = true; } catch { }
                BreakExternalLinks(result);
                result.SaveAs(output, 51);
                result.Close(false);
                result = null;
                return output;
            }
            catch
            {
                if (result != null) try { result.Close(false); } catch { }
                if (File.Exists(output)) try { File.Delete(output); } catch { }
                throw;
            }
            }
        }

        private static string NewOutputPath(string sourcePath, string exportWorkbookName)
        {
            string folder = Path.Combine(Path.GetDirectoryName(sourcePath), "提项目");
            Directory.CreateDirectory(folder);
            string stem = HydroExportPolicy.FileStem(DateTime.Now, exportWorkbookName);
            string path = Path.Combine(folder, stem + ".xlsx");
            for (int n = 2; File.Exists(path); n++) path = Path.Combine(folder, stem + "_" + n.ToString(CultureInfo.InvariantCulture) + ".xlsx");
            return path;
        }

        private static string SafeSheetName(string name, dynamic workbook)
        {
            string safe = name;
            foreach (char c in new[] { '[', ']', ':', '*', '?', '/', '\\' }) safe = safe.Replace(c, '_');
            if (safe.Length > 31) safe = safe.Substring(0, 31);
            string baseName = safe; int n = 2;
            while (HasSheet(workbook, safe))
            {
                string suffix = "_" + n.ToString(CultureInfo.InvariantCulture);
                safe = baseName.Substring(0, Math.Min(baseName.Length, 31 - suffix.Length)) + suffix; n++;
            }
            return safe;
        }

        private static bool HasSheet(dynamic workbook, string name)
        {
            foreach (dynamic sheet in workbook.Worksheets)
                if (string.Equals(Convert.ToString(sheet.Name, CultureInfo.InvariantCulture), name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        private static void BreakExternalLinks(dynamic workbook)
        {
            foreach (int kind in new[] { 1, 2 })
            {
                object links = null;
                try { links = workbook.LinkSources(kind); } catch { }
                if (links == null) continue;
                var array = links as Array;
                if (array != null)
                {
                    foreach (object link in array) try { workbook.BreakLink(link, kind); } catch { }
                }
                else
                {
                    try { workbook.BreakLink(links, kind); } catch { }
                }
            }
        }

        private static void DeleteComments(dynamic sheet, dynamic range)
        {
            try { range.ClearComments(); } catch { }
            try
            {
                dynamic threaded = sheet.CommentsThreaded;
                for (int i = (int)threaded.Count; i >= 1; i--) threaded[i].Delete();
            }
            catch { }
            DeleteObjects(sheet);
        }

        private static void DeleteObjects(dynamic sheet)
        {
            for (int i = (int)sheet.Shapes.Count; i >= 1; i--) try { sheet.Shapes[i].Delete(); } catch { }
            try { sheet.DrawingObjects().Delete(); } catch { }
            try { sheet.Cells.Hyperlinks.Delete(); } catch { }
            try { if ((bool)sheet.AutoFilterMode) sheet.AutoFilterMode = false; } catch { }
        }

        private static void FormatRange(dynamic sheet, dynamic range)
        {
            range.Font.Name = "宋体";
            range.Font.Name = "Times New Roman";
            range.Font.Size = 10.5;
            range.Font.Color = 0;
            object titleValues = range.Resize[Math.Min(6, (int)range.Rows.Count), (int)range.Columns.Count].Value2;
            var matrix = titleValues as object[,];
            for (int r = 1; r <= Math.Min(6, (int)range.Rows.Count); r++)
            {
                var values = new List<string>();
                for (int c = 1; c <= (int)range.Columns.Count; c++)
                    values.Add(Convert.ToString(matrix == null ? titleValues : matrix[r, c], CultureInfo.InvariantCulture));
                if (HydroExportPolicy.IsTitleRow(values.ToArray(), r - 1))
                {
                    dynamic title = range.Rows[r];
                    title.Font.Name = "黑体";
                    title.Font.Size = 12;
                }
            }
            range.VerticalAlignment = -4108;
            range.WrapText = true;
            range.Columns.AutoFit();
            for (int c = 1; c <= (int)range.Columns.Count; c++)
            {
                dynamic col = range.Columns[c];
                if ((double)col.ColumnWidth > 60) col.ColumnWidth = 60;
            }
            range.Rows.AutoFit();
            int rowCount = (int)range.Rows.Count;
            Func<int, int, double?> readHeight = (first, last) =>
            {
                object height = sheet.Range[sheet.Cells[first, 1], sheet.Cells[last, 1]].EntireRow.RowHeight;
                return height == null || height == DBNull.Value ? (double?)null : Convert.ToDouble(height, CultureInfo.InvariantCulture);
            };
            Action<int, int, double> writeHeight = (first, last, height) =>
                sheet.Range[sheet.Cells[first, 1], sheet.Cells[last, 1]].EntireRow.RowHeight = height;
            HydroRowHeights.Pad(rowCount, readHeight, writeHeight);
        }
    }
}
