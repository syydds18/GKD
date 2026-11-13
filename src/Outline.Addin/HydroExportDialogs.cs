using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Collections.Specialized;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Outline.Addin
{
    internal sealed class HydroConfirmForm : Form
    {
        internal HydroConfirmForm(List<HydroSheetChoice> selected)
        {
            Ui.Base(this, new Size(800, 540), "水电工程提项目 — 确认");
            MinimumSize = new Size(640, 420);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            layout.Controls.Add(new Label { AutoSize = true, Dock = DockStyle.Fill, MaximumSize = new Size(720, 0), Text = "已选择 " + selected.Count + " 张工作表。仅在导出副本中清理公式、外链、对象、图片和批注，源工作簿不会修改。" }, 0, 0);
            var list = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, ScrollBars = ScrollBars.Both };
            list.Columns.Add("sheet", "工作表");
            list.Columns.Add("range", "打印范围");
            foreach (var choice in selected) list.Rows.Add(choice.SheetName, choice.PrintRange);
            layout.Controls.Add(list, 0, 1);
            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, Padding = new Padding(0, 8, 0, 0) };
            var cancel = Ui.Button("取消"); cancel.DialogResult = DialogResult.Cancel;
            var confirm = Ui.Button("确认导出", true); confirm.DialogResult = DialogResult.OK;
            footer.Controls.Add(cancel); footer.Controls.Add(confirm);
            layout.Controls.Add(footer, 0, 2);
            Controls.Add(layout); AcceptButton = confirm; CancelButton = cancel;
            HydroWindowBounds.Fit(this);
        }
    }

    internal sealed class HydroProgressForm : Form
    {
        private readonly Label message = new Label { Dock = DockStyle.Top, Height = 48 };
        private readonly ProgressBar bar = new ProgressBar { Dock = DockStyle.Top, Height = 24, Minimum = 0 };
        internal HydroProgressForm(int total)
        {
            Ui.Base(this, new Size(520, 140), "水电工程 — 正在导出");
            ControlBox = false; MaximizeBox = false; MinimizeBox = false;
            bar.Maximum = Math.Max(1, total);
            Controls.Add(bar); Controls.Add(message);
        }
        internal void Report(int completed, string stage)
        {
            message.Text = stage;
            bar.Value = Math.Max(0, Math.Min(bar.Maximum, completed));
            // Paint this window only; do not pump workbook commands or use worker-thread COM.
            Refresh();
        }
    }

    internal sealed class HydroSuccessForm : Form
    {
        private readonly string outputPath;

        internal HydroSuccessForm(string output)
        {
            outputPath = output;
            Ui.Base(this, new Size(720, 250), "水电工程提项目 — 导出完成");
            MinimumSize = new Size(560, 220);
            MaximizeBox = false;
            MinimizeBox = false;

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            layout.Controls.Add(new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Text = "导出完成，源工作簿未修改。"
            }, 0, 0);
            layout.Controls.Add(new TextBox
            {
                ReadOnly = true,
                Multiline = true,
                BorderStyle = BorderStyle.None,
                BackColor = BackColor,
                Dock = DockStyle.Fill,
                Text = outputPath,
                Margin = new Padding(0, 12, 0, 0),
                TabStop = false
            }, 0, 1);

            // Keep the visual order explicit: open file, folder, file-copy, OK.
            var footer = new FlowLayoutPanel
            {
                Dock = DockStyle.None,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 8, 0, 0),
                Anchor = AnchorStyles.None,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0)
            };
            var openFile = Ui.Button("打开导表");
            openFile.BackColor = Color.FromArgb(255, 248, 214);
            openFile.FlatAppearance.BorderColor = Color.FromArgb(231, 205, 122);
            openFile.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 240, 179);
            openFile.FlatAppearance.MouseDownBackColor = Color.FromArgb(250, 226, 145);
            openFile.Click += (s, e) => { OpenOutputFile(); Close(); };
            var openFolder = Ui.Button("打开所在文件夹");
            openFolder.Click += (s, e) => { OpenContainingFolder(); Close(); };
            var copyFile = Ui.Button("复制该文件");
            copyFile.Click += (s, e) => { CopyFileToClipboard(); Close(); };
            var ok = Ui.Button("确定", true);
            ok.DialogResult = DialogResult.OK;
            openFile.Margin = new Padding(4, 4, 6, 4);
            openFolder.Margin = new Padding(6, 4, 6, 4);
            copyFile.Margin = new Padding(6, 4, 6, 4);
            ok.Margin = new Padding(6, 4, 4, 4);
            foreach (var button in new[] { openFile, openFolder, copyFile, ok }) button.Padding = new Padding(8, 3, 8, 3);
            footer.Controls.Add(openFile);
            footer.Controls.Add(openFolder);
            footer.Controls.Add(copyFile);
            footer.Controls.Add(ok);
            layout.Controls.Add(footer, 0, 2);
            Controls.Add(layout);
            AcceptButton = ok;
            HydroWindowBounds.Fit(this);
        }

        private bool TryGetOutput(out string path)
        {
            path = outputPath;
            if (string.IsNullOrWhiteSpace(path))
            {
                MessageBox.Show(this, "没有可用的导出文件路径。", "提项目", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            try
            {
                if (!File.Exists(path))
                {
                    MessageBox.Show(this, "导出文件不存在，无法执行此操作：\r\n" + path, "提项目", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                path = Path.GetFullPath(path);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法检查导出文件：" + ex.Message, "提项目", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void OpenContainingFolder()
        {
            string path;
            if (!TryGetOutput(out path)) return;
            try
            {
                string argument = "/select,\"" + path.Replace("\"", string.Empty) + "\"";
                Process.Start(new ProcessStartInfo("explorer.exe", argument) { UseShellExecute = false });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法打开所在文件夹：" + ex.Message, "提项目", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenOutputFile()
        {
            string path;
            if (!TryGetOutput(out path)) return;
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法打开导出文件：" + ex.Message, "提项目", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CopyFileToClipboard()
        {
            string path;
            if (!TryGetOutput(out path)) return;
            try
            {
                var files = new StringCollection();
                files.Add(path);
                Exception lastError = null;
                var thread = new System.Threading.Thread(() =>
                {
                    try
                    {
                        for (int attempt = 0; attempt < 3; attempt++)
                        {
                            try
                            {
                                var data = new DataObject();
                                data.SetFileDropList(files);
                                Clipboard.SetDataObject(data, true, 10, 100);
                                lastError = null;
                                break;
                            }
                            catch (ExternalException ex)
                            {
                                lastError = ex;
                                System.Threading.Thread.Sleep(40 * (attempt + 1));
                            }
                        }
                    }
                    catch (Exception ex) { lastError = ex; }
                });
                thread.IsBackground = true;
                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                thread.Start();
                if (!thread.Join(4000)) lastError = new TimeoutException("剪贴板操作超时。");
                if (lastError != null) throw lastError;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法复制导出文件：" + ex.Message, "提项目", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    internal static class HydroWindowBounds
    {
        internal static void Fit(Form form)
        {
            form.Shown += (s, e) =>
            {
                Rectangle area = Screen.FromControl(form).WorkingArea;
                form.MinimumSize = new Size(Math.Min(form.MinimumSize.Width, area.Width), Math.Min(form.MinimumSize.Height, area.Height));
                form.Size = new Size(Math.Min(form.Width, area.Width), Math.Min(form.Height, area.Height));
                form.Location = new Point(area.Left + (area.Width - form.Width) / 2, area.Top + (area.Height - form.Height) / 2);
            };
        }
    }
}
