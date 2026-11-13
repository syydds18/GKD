using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Outline.Setup
{
    internal sealed class MainForm : Form
    {
        private readonly ComboBox targets = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, IntegralHeight = false, DropDownHeight = 220 };
        private readonly Button install = new Button { Text = "一键安装", AutoSize = true, MinimumSize = new Size(128, 42) };
        private readonly Button uninstall = new Button { Text = "卸载加载项", AutoSize = true, MinimumSize = new Size(128, 42) };
        private readonly Button forceUninstall = new Button { Text = "强制关闭 Excel 并卸载…", AutoSize = true, MinimumSize = new Size(190, 42), Enabled = false };
        private bool forceAvailable;
        private readonly Button refresh = new Button { Text = "重新检测", AutoSize = true, MinimumSize = new Size(100, 34) };
        private readonly Button browse = new Button { Text = "选择 EXCEL.EXE…", AutoSize = true, MinimumSize = new Size(152, 34) };
        private readonly TextBox log = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.FromArgb(247, 250, 248), MinimumSize = new Size(0, 110) };
        private readonly Label state = new Label { AutoSize = true, Dock = DockStyle.Fill, ForeColor = Color.FromArgb(20, 111, 70), Padding = new Padding(0, 10, 0, 8) };
        private readonly bool uninstallMode;
        private bool busy;
        internal MainForm(bool uninstallMode)
        {
            this.uninstallMode = uninstallMode;
            Text = "搞快点 · 安装与卸载";
            Font = new Font("Microsoft YaHei UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            BackColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(800, 655);
            MinimumSize = new Size(700, 650);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 9, Padding = new Padding(26, 22, 26, 20) };
            for (int i = 0; i < 9; i++) layout.RowStyles.Add(new RowStyle(i == 7 ? SizeType.Percent : SizeType.AutoSize, i == 7 ? 100 : 0));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.Controls.Add(new Label { Text = "搞快点", Font = new Font(Font.FontFamily, 19, FontStyle.Bold), AutoSize = true, Margin = new Padding(0, 0, 0, 6) }, 0, 0);
            layout.Controls.Add(new Label { Text = "原生 Excel 加载项  /  v1.4.0  /  仅当前 Windows 用户，无需管理员", AutoSize = true, ForeColor = Color.FromArgb(95, 103, 98), Margin = new Padding(0, 0, 0, 12) }, 0, 1);
            layout.Controls.Add(state, 0, 2);
            var detection = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3, Margin = new Padding(0, 0, 0, 12) };
            detection.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); detection.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            detection.Controls.Add(new Label { Text = "目标 Excel（读取 EXCEL.EXE 的 PE 位数，不使用操作系统位数推断）", AutoSize = true, Margin = new Padding(0, 4, 0, 5) }, 0, 0);
            detection.SetColumnSpan(detection.GetControlFromPosition(0, 0), 2);
            detection.Controls.Add(targets, 0, 1); detection.Controls.Add(browse, 1, 1);
            detection.Controls.Add(refresh, 0, 2);
            layout.Controls.Add(detection, 0, 3);
            var path = new TextBox { ReadOnly = true, Text = Installer.InstallRoot, Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
            layout.Controls.Add(new Label { Text = "安装位置（固定在当前用户本地目录）", AutoSize = true, Margin = new Padding(0, 4, 0, 5) }, 0, 4);
            layout.Controls.Add(path, 0, 5);
            layout.Controls.Add(new Label { Text = "安装前请保存并关闭 Excel；卸载先尝试正常关闭。超时后才可单独选择强制关闭（未保存内容可能丢失）。未签名 XLL 可能被企业策略拦截；本程序不修改宏设置、信任中心或受信任位置。", AutoSize = true, MaximumSize = new Size(730, 0), ForeColor = Color.FromArgb(105, 95, 58), Margin = new Padding(0, 16, 0, 12) }, 0, 6);
            layout.Controls.Add(log, 0, 7);
            var actions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = true, Padding = new Padding(0, 12, 0, 0) };
            install.BackColor = Color.FromArgb(24, 123, 77); install.ForeColor = Color.White; install.FlatStyle = FlatStyle.Flat; install.FlatAppearance.BorderSize = 0;
            var help = new Button { Text = "使用帮助", AutoSize = true, MinimumSize = new Size(100, 42) };
            actions.Controls.Add(install); actions.Controls.Add(uninstall); actions.Controls.Add(forceUninstall); actions.Controls.Add(help);
            layout.Controls.Add(actions, 0, 8);
            Controls.Add(layout);
            install.Click += (s, e) => RunAction(false);
            uninstall.Click += (s, e) => RunAction(true);
            forceUninstall.Click += (s, e) => RunAction(true, true);
            refresh.Click += (s, e) => Discover();
            browse.Click += (s, e) => Browse();
            help.Click += (s, e) => MessageBox.Show(this,
                "1. 解压完整安装包，保持安装程序与 payload 文件夹相邻。\r\n\r\n2. 选择要使用的 Excel 版本与位数；未检测到时可浏览 EXCEL.EXE。\r\n\r\n3. 点击“一键安装”，然后启动 Excel，在“搞快点”选项卡使用功能。若未显示，请联系 IT 检查 XLL 签名与组织策略。\r\n\r\n4. 可在本窗口或 Windows“应用/程序和功能”中卸载。安装前请保存并关闭 Excel。卸载会先请求正常关闭；超时后可另行确认强制关闭当前用户、当前会话中的 Excel，未保存内容可能永久丢失。\r\n\r\n运行要求：Windows 桌面版 Excel 2010+、.NET Framework 4.8；支持 x86/x64。", "搞快点使用帮助", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Shown += (s, e) => { Discover(); if (uninstallMode) { uninstall.Focus(); Append("已进入卸载模式。请点击“卸载加载项”确认操作。"); } };
            FormClosing += (s, e) => { if (busy) e.Cancel = true; };
        }
        private void Discover()
        {
            try
            {
                targets.Items.Clear();
                foreach (var target in ExcelDiscovery.Find(Append)) targets.Items.Add(target);
                if (targets.Items.Count > 0) targets.SelectedIndex = 0;
                Append(targets.Items.Count == 0 ? "未自动发现 Excel。可点击“选择 EXCEL.EXE…”手动定位真实程序。" : "已检测到 " + targets.Items.Count + " 个 Excel 安装，请核对版本与位数。");
                UpdateState();
            }
            catch (Exception ex) { Append("检测失败：" + ex.Message); }
        }
        private void Browse()
        {
            using (var dialog = new OpenFileDialog { Title = "选择 Excel 程序", Filter = "Excel 程序 (EXCEL.EXE)|EXCEL.EXE", CheckFileExists = true, Multiselect = false })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try { var target = ExcelDiscovery.Inspect(dialog.FileName); targets.Items.Add(target); targets.SelectedIndex = targets.Items.Count - 1; Append("已核对 PE 架构：" + target.Architecture); }
                catch (Exception ex) { MessageBox.Show(this, ex.Message, "无法选择", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            }
        }
        private async void RunAction(bool remove, bool force = false)
        {
            if (busy || (force && (!remove || !forceAvailable))) return;
            var target = targets.SelectedItem as ExcelTarget;
            if (!remove && target == null) { MessageBox.Show(this, "请先检测或选择 EXCEL.EXE。", "尚未选择 Excel"); return; }
            if (remove && MessageBox.Show(this, "卸载当前用户的搞快点加载项？\r\n如 Excel 正在运行，将另行询问是否关闭。", "确认卸载", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            try
            {
                busy = true; install.Enabled = uninstall.Enabled = forceUninstall.Enabled = browse.Enabled = refresh.Enabled = targets.Enabled = false;
                if (remove)
                {
                    bool closed;
                    if (force)
                    {
                        // Capture before asking: consent covers exactly these verified process handles.
                        using (var snapshot = ExcelShutdown.CaptureForForce())
                        {
                            closed = await ExcelCloseGate.ForceClosed(ExcelShutdown.Count,
                                count => MessageBox.Show(this, "危险：正常关闭未完成。是否强制结束当前用户、当前 Windows 会话的 " + snapshot.Count + " 个 Excel 进程并继续卸载？\r\n\r\n这会立即关闭这些进程中的全部 Excel 工作簿（不只是后台实例），不会显示保存提示。未保存的内容可能永久丢失！\r\n\r\n请先保存所有文件。其他用户、其他会话和确认后新启动的 Excel 不会被强制结束。无法核对身份或权限不足时停止，不会自动提权。\r\n\r\n选择“否”取消，不会结束任何进程。", "确认丢失风险：强制结束 Excel？", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes,
                                () => snapshot.ForceClose(Append), () => Task.Delay(500));
                        }
                    }
                    else
                    {
                        forceAvailable = false;
                        var result = await ExcelCloseGate.CloseNormally(ExcelShutdown.Count,
                            count => MessageBox.Show(this, "检测到当前用户、当前会话的 " + count + " 个 Excel 进程。是否请求正常关闭后继续卸载？\r\n\r\n未保存的文件将由 Excel 询问是否保存，请先处理保存提示。此操作不会自动放弃更改或强制结束进程。\r\n\r\n选择“否”取消本次卸载。", "关闭 Excel 后继续？", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes,
                            () => ExcelShutdown.RequestClose(Append), () => Task.Delay(500));
                        forceAvailable = result == ExcelCloseResult.StillRunning;
                        closed = result == ExcelCloseResult.Closed;
                        if (forceAvailable) Append("正常关闭等待超时。请处理保存提示或后台实例后重试；如确认允许丢失全部未保存内容，可单独点击“强制关闭 Excel 并卸载…”。不会自动强制结束。");
                    }
                    if (!closed) { Append("Excel 未退出或操作已取消，没有执行卸载。"); return; }
                }
                UseWaitCursor = true;
                if (remove) Installer.Uninstall(Append); else Installer.Install(target, Append);
                forceAvailable = false;
                MessageBox.Show(this, remove ? "卸载操作已完成。详情请看状态记录。" : "安装完成。请自行启动 Excel。若被组织安全策略拦截，请联系 IT。", "操作完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { Append("操作停止：" + ex.Message); MessageBox.Show(this, ex.Message, "操作未完成", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { busy = false; UseWaitCursor = false; browse.Enabled = refresh.Enabled = targets.Enabled = true; UpdateState(); }
        }
        private void UpdateState()
        {
            bool exists = File.Exists(Installer.StatePath);
            try
            {
                if (exists)
                {
                    var saved = StateFile.Load(Installer.StatePath);
                    state.Text = (saved.Complete ? "已安装" : "安装未完成，可卸载恢复") + "  ·  Excel " + saved.OfficeVersion + "  ·  " + saved.Architecture;
                }
                else state.Text = "尚未安装 · v1.4.0";
            }
            catch (Exception ex) { state.Text = "安装记录无法读取，请保留目录并联系维护人员"; Append(ex.Message); }
            install.Enabled = !exists && Directory.Exists(Installer.PayloadRoot);
            uninstall.Enabled = exists;
            forceUninstall.Enabled = exists && forceAvailable && !busy;
            if (!Directory.Exists(Installer.PayloadRoot) && !exists) Append("缺少相邻的 payload 文件夹。请使用完整解压的发行包安装。");
        }
        private void Append(string text) { log.AppendText(text + Environment.NewLine); log.SelectionStart = log.TextLength; log.ScrollToCaret(); log.Refresh(); }
    }
}
