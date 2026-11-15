using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Outline.Addin
{
    internal sealed class GkdAiSettingsForm : Form
    {
        private readonly ListBox profiles = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, DisplayMember = "Caption" };
        private readonly TableLayoutPanel settingsPanel = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Padding = new Padding(4) };
        private readonly TextBox profileName = Input();
        private readonly ComboBox provider = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox endpoint = Input();
        private readonly TextBox model = Input();
        private readonly TextBox apiKey = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, UseSystemPasswordChar = true };
        private readonly Button showApiKey = new Button { Text = "", AccessibleName = "显示 API Key", AutoSize = false, Width = 28, Dock = DockStyle.Right, Margin = Padding.Empty, Padding = Padding.Empty, FlatStyle = FlatStyle.Flat, TabStop = false, BackColor = Color.White, UseVisualStyleBackColor = false };
        private readonly NumericUpDown timeout = new NumericUpDown { Dock = DockStyle.Left, Minimum = 1, Maximum = 600, Value = 120, Width = 110 };
        private readonly TextBox promptInput = Input();
        private readonly TextBox routeInput = Input();
        private readonly TextBox answerOutput = Input();
        private readonly GroupBox difyFields = new GroupBox { Text = "Dify 工作流字段映射", Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(10, 20, 10, 8) };
        private readonly Label draftStatus = new Label { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 6, 0, 8) };
        private readonly TextBox testStatus = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = Color.White, Text = "尚未测试。测试基于当前设置（包含未保存的编辑），不会自动保存；结果仅代表该次单轮请求。" };
        private readonly Button create = Ui.Button("新建");
        private readonly Button save = Ui.Button("保存配置", true);
        private readonly Button delete = Ui.Button("删除配置");
        private readonly Button makeDefault = Ui.Button("设为默认");
        private readonly Button connectionTest = Ui.Button("连接 / 鉴权测试");
        private readonly Button tableTest = Ui.Button("表格能力测试");
        private readonly Button cancelTest = Ui.Button("取消测试");
        private readonly Button loadZhiqingTest = Ui.Button("填入已验证智擎测试配置");
        private readonly Button help = Ui.Button("★ 使用帮助与示例");
        private readonly Button openLogs = Ui.Button("打开诊断日志");
        private GkdAiSettingsEditor editor;
        private string currentId;
        private bool dirty;
        private bool loading;
        private bool closing;
        private CancellationTokenSource cancellation;

        internal GkdAiSettingsForm()
        {
            Ui.Base(this, new Size(1080, 790), "AI 模型设置 — GKD 公式");
            MinimumSize = new Size(960, 720);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            var header = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 0, 0, 8) };
            header.Controls.Add(new Label { Text = "设置默认模型后，在工作表中使用 GKD 公式。", AutoSize = true, Margin = new Padding(0, 14, 12, 4) });
            header.Controls.Add(help);
            layout.Controls.Add(header, 0, 0);
            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 242));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Padding = new Padding(0, 0, 16, 0) };
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 170));
            left.Controls.Add(new Label { Text = "已保存模型配置（★ 为默认）", Dock = DockStyle.Fill }, 0, 0);
            left.Controls.Add(profiles, 0, 1);
            var profileActions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = true };
            profileActions.Controls.AddRange(new Control[] { create, save, delete, makeDefault });
            left.Controls.Add(profileActions, 0, 2);
            body.Controls.Add(left, 0, 0);
            var scroller = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 122));
            settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            provider.Items.AddRange(new object[] { "OpenAI 兼容接口", "Dify 工作流" });
            provider.SelectedIndex = 0;
            AddField(settingsPanel, "配置名称", profileName);
            AddField(settingsPanel, "接口类型", provider);
            AddField(settingsPanel, "完整请求 URL", endpoint);
            AddField(settingsPanel, "模型名称 / 路由（可选）", model);
            var keyRow = new Panel { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Height = apiKey.PreferredHeight + 10, Padding = new Padding(5, 4, 2, 3), Margin = Padding.Empty };
            showApiKey.FlatAppearance.BorderSize = 0;
            apiKey.Margin = Padding.Empty;
            keyRow.Controls.Add(apiKey);
            keyRow.Controls.Add(showApiKey);
            showApiKey.Paint += PaintApiKeyEye;
            AddField(settingsPanel, "API Key", keyRow);
            AddField(settingsPanel, "超时（秒）", timeout);
            var mapping = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
            mapping.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108));
            mapping.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            AddField(mapping, "提示词输入", promptInput);
            AddField(mapping, "模型路由输入（可选）", routeInput);
            AddField(mapping, "答案输出", answerOutput);
            var preset = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
            loadZhiqingTest.Margin = new Padding(0, 4, 0, 4);
            preset.Controls.Add(loadZhiqingTest);
            var difyLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };
            difyLayout.Controls.Add(mapping, 0, 0);
            difyLayout.Controls.Add(preset, 0, 1);
            difyFields.Controls.Add(difyLayout);
            AddWide(settingsPanel, difyFields);

            AddWide(settingsPanel, draftStatus);
            scroller.Controls.Add(settingsPanel);
            body.Controls.Add(scroller, 1, 0);
            layout.Controls.Add(body, 0, 1);
            var testing = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
            testing.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            testing.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var testActions = new FlowLayoutPanel { Dock = DockStyle.Fill };
            testActions.Controls.AddRange(new Control[] { connectionTest, tableTest, cancelTest, openLogs });
            testing.Controls.Add(testActions, 0, 0);
            testing.Controls.Add(testStatus, 0, 1);
            layout.Controls.Add(testing, 0, 2);
            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            var close = Ui.Button("关闭"); close.Click += (s, e) => { SetApiKeyVisible(false); Close(); };
            footer.Controls.Add(close);
            layout.Controls.Add(footer, 0, 3);
            Controls.Add(layout);
            CancelButton = close;
            help.Click += (s, e) => { using (var dialog = new GkdAiHelpForm()) dialog.ShowDialog(this); };
            openLogs.Click += (s, e) => OpenDiagnosticsDirectory();
            showApiKey.Click += (s, e) => SetApiKeyVisible(apiKey.UseSystemPasswordChar);
            create.Click += (s, e) => NewProfile();
            save.Click += (s, e) => SaveCurrent();
            delete.Click += (s, e) => DeleteCurrent();
            makeDefault.Click += (s, e) => SetCurrentDefault();
            connectionTest.Click += async (s, e) => await RunTestAsync(false);
            tableTest.Click += async (s, e) => await RunTestAsync(true);
            cancelTest.Click += (s, e) => CancelTest();
            loadZhiqingTest.Click += (s, e) => { try { LoadVerifiedZhiqingTest(); } catch (Exception) { MessageBox.Show(this, "测试配置读取失败，请检查安装包是否完整。", "AI 模型设置", MessageBoxButtons.OK, MessageBoxIcon.Warning); } };
            profiles.SelectedIndexChanged += (s, e) => SelectProfile();
            foreach (var input in new Control[] { profileName, endpoint, model, apiKey, promptInput, routeInput, answerOutput })
                input.TextChanged += (s, e) => MarkDirty();
            timeout.ValueChanged += (s, e) => MarkDirty();
            provider.SelectedIndexChanged += (s, e) => { difyFields.Visible = provider.SelectedIndex == 1; MarkDirty(); };
            bool loadFailed = false;
            try { editor = new GkdAiSettingsEditor(GkdAiStore.Load()); }
            catch (Exception) { editor = new GkdAiSettingsEditor(new GkdAiSettings()); loadFailed = true; }
            var loaded = editor.Snapshot();
            string initialId = loaded.DefaultProfileId;
            if (string.IsNullOrEmpty(initialId) && loaded.Profiles.Count != 0) initialId = loaded.Profiles[0].Id;
            RefreshProfiles(initialId);
            if (profiles.SelectedItem != null) ShowProfile(editor.Find(((ProfileChoice)profiles.SelectedItem).Id));
            else ShowProfile(editor.CreateProfile());
            SetBusy(false);
            if (loadFailed)
                testStatus.Text = "加载配置失败，可能是文件损坏或密钥无法解密。原文件尚未修改；保存新配置会替换它。请先备份或修复后重新打开。本次尚未测试。";
        }

        private void LoadVerifiedZhiqingTest()
        {
            if (!MayDiscard()) return;
            loading = true;
            try
            {
                currentId = editor.CreateProfile().Id;
                profileName.Text = "搞快点智擎工作流";
                provider.SelectedIndex = 1;
                endpoint.Text = "https://aibuilder.hdec.com/v1/workflows/run";
                model.Text = "hdec235";
                string configDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelOutlineAddin", "1.0");
                string keyPath = Path.Combine(configDirectory, "gkd-workflow.key");
                if (!File.Exists(keyPath)) keyPath = Path.Combine(configDirectory, "gkd-test-workflow.key");
                apiKey.Text = File.Exists(keyPath) ? File.ReadAllText(keyPath).Trim() : "";
                promptInput.Text = "question";
                routeInput.Text = "route";
                answerOutput.Text = "answer";
                difyFields.Visible = true;
                SetApiKeyVisible(false);
                dirty = true;
            }
            finally { loading = false; }
            UpdateDraftStatus();
            testStatus.Text = "已填入搞快点智擎工作流参数；若密钥为空，请粘贴该工作流应用自己的 API Key。请执行连接测试；此预设不会自动保存。";
        }
        private void OpenDiagnosticsDirectory()
        {
            try
            {
                string directory = GkdAiDiagnostics.LogDirectory;
                Directory.CreateDirectory(directory);
                Process.Start(new ProcessStartInfo { FileName = directory, UseShellExecute = true });
            }
            catch (Exception)
            {
                MessageBox.Show(this, "无法打开诊断日志目录，请检查当前用户目录的访问权限及资源管理器。", "AI 模型设置", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private static TextBox Input() { return new TextBox { Dock = DockStyle.Fill }; }
        private void SetApiKeyVisible(bool visible)
        {
            apiKey.UseSystemPasswordChar = !visible;
            showApiKey.AccessibleName = visible ? "隐藏 API Key" : "显示 API Key";
            showApiKey.Invalidate();
        }
        private void PaintApiKeyEye(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            float scale = Math.Min(showApiKey.ClientSize.Width / 24f, showApiKey.ClientSize.Height / 20f);
            float x = showApiKey.ClientSize.Width / 2f, y = showApiKey.ClientSize.Height / 2f;
            using (var pen = new Pen(showApiKey.Enabled ? Color.DimGray : SystemColors.GrayText, 1.4f * scale))
            using (var path = new System.Drawing.Drawing2D.GraphicsPath())
            {
                path.AddBezier(x - 8 * scale, y, x - 4 * scale, y - 7 * scale, x + 4 * scale, y - 7 * scale, x + 8 * scale, y);
                path.AddBezier(x + 8 * scale, y, x + 4 * scale, y + 7 * scale, x - 4 * scale, y + 7 * scale, x - 8 * scale, y);
                g.DrawPath(pen, path);
                g.DrawEllipse(pen, x - 2 * scale, y - 2 * scale, 4 * scale, 4 * scale);
                if (apiKey.UseSystemPasswordChar) g.DrawLine(pen, x - 7 * scale, y + 6 * scale, x + 7 * scale, y - 6 * scale);
            }
        }
        private static void AddField(TableLayoutPanel table, string caption, Control input)
        {
            int row = table.RowCount++;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(new Label { Text = caption, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 8, 6, 8) }, 0, row);
            input.Margin = new Padding(0, 5, 0, 5);
            table.Controls.Add(input, 1, row);
        }
        private static void AddWide(TableLayoutPanel table, Control control)
        {
            int row = table.RowCount++;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(control, 0, row); table.SetColumnSpan(control, 2);
        }
        private sealed class ProfileChoice
        {
            public string Id { get; set; }
            public string Caption { get; set; }
        }
        private void RefreshProfiles(string selectedId)
        {
            loading = true;
            try
            {
                profiles.Items.Clear();
                var snapshot = editor.Snapshot();
                foreach (var profile in snapshot.Profiles)
                {
                    int index = profiles.Items.Add(new ProfileChoice { Id = profile.Id, Caption = (profile.Id == snapshot.DefaultProfileId ? "★ " : "") + profile.Name });
                    if (profile.Id == selectedId) profiles.SelectedIndex = index;
                }
            }
            finally { loading = false; }
        }
        private void SelectProfile()
        {
            if (loading || cancellation != null || profiles.SelectedItem == null) return;
            string selectedId = ((ProfileChoice)profiles.SelectedItem).Id;
            if (selectedId == currentId) return;
            if (!MayDiscard()) { RefreshProfiles(currentId); return; }
            ShowProfile(editor.Find(selectedId));
        }
        private void ShowProfile(GkdAiProfile profile)
        {
            loading = true;
            try
            {
                SetApiKeyVisible(false);
                currentId = profile.Id; profileName.Text = profile.Name; provider.SelectedIndex = profile.Provider == "dify" ? 1 : 0;
                endpoint.Text = profile.Endpoint; model.Text = profile.Model; apiKey.Text = profile.ApiKey;
                timeout.Value = Math.Max(timeout.Minimum, Math.Min(timeout.Maximum, profile.TimeoutSeconds));
                promptInput.Text = profile.PromptInput; routeInput.Text = profile.RouteInput; answerOutput.Text = profile.AnswerOutput;
                difyFields.Visible = provider.SelectedIndex == 1;
                dirty = false;
                testStatus.Text = "当前配置尚未测试。基于当前设置（含未保存内容）发送少量合成测试数据，不读取工作簿。";
            }
            finally { loading = false; }
            UpdateDraftStatus();
        }
        private GkdAiProfile ReadDraft()
        {
            return new GkdAiProfile
            {
                Id = currentId, Name = profileName.Text.Trim(), Provider = provider.SelectedIndex == 1 ? "dify" : "openai",
                Endpoint = endpoint.Text.Trim(), Model = model.Text.Trim(), ApiKey = apiKey.Text.Trim(),
                TimeoutSeconds = Decimal.ToInt32(timeout.Value), PromptInput = promptInput.Text.Trim(),
                RouteInput = routeInput.Text.Trim(), AnswerOutput = answerOutput.Text.Trim()
            };
        }
        private void MarkDirty()
        {
            if (loading) return;
            dirty = true; UpdateDraftStatus();
            testStatus.Text = "当前设置已修改，先前结果不再适用。基于当前设置测试不会自动保存。";
        }
        private void UpdateDraftStatus()
        {
            draftStatus.Text = dirty || editor.Find(currentId) == null ? "有未保存的设置；测试使用当前内容。" : "配置已保存；测试结果不会写入配置。";
            bool saved = editor.Find(currentId) != null;
            delete.Enabled = saved && cancellation == null;
            makeDefault.Enabled = saved && cancellation == null;
        }
        private bool MayDiscard()
        {
            return !dirty || MessageBox.Show(this, "当前设置尚未保存。是否放弃这些修改？", "AI 模型设置", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }
        private void NewProfile()
        {
            if (!MayDiscard()) return;
            RefreshProfiles(null); ShowProfile(editor.CreateProfile());
        }
        private void SaveCurrent()
        {
            try
            {
                var candidate = new GkdAiSettingsEditor(editor.Snapshot());
                candidate.SaveProfile(ReadDraft());
                GkdAiStore.Save(candidate.Snapshot()); editor = candidate;
                GkdAiFunctions.ConfigurationChanged();
                dirty = false; RefreshProfiles(currentId); UpdateDraftStatus();
            }
            catch (ArgumentException ex) { MessageBox.Show(this, "保存失败：" + Redact(ex.Message, ReadDraft()), "AI 模型设置", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (Exception) { MessageBox.Show(this, "保存失败。请检查配置目录是否可写。", "AI 模型设置", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        private void DeleteCurrent()
        {
            if (editor.Find(currentId) == null) return;
            if (editor.Snapshot().Profiles.Count == 1)
            {
                MessageBox.Show(this, "至少保留一个已保存配置。请先新建并保存替代配置，再删除本项。", "AI 模型设置");
                return;
            }
            if (MessageBox.Show(this, "删除当前已保存配置？未保存的编辑也将丢弃。若它是默认配置，将改用剩余首项。", "AI 模型设置", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            try
            {
                var candidate = new GkdAiSettingsEditor(editor.Snapshot()); candidate.DeleteProfile(currentId);
                GkdAiStore.Save(candidate.Snapshot()); editor = candidate;
                GkdAiFunctions.ConfigurationChanged();
                string next = editor.Snapshot().DefaultProfileId;
                RefreshProfiles(next); ShowProfile(editor.Find(next) ?? editor.CreateProfile());
            }
            catch (Exception) { MessageBox.Show(this, "删除配置失败，原配置已保留。", "AI 模型设置"); }
        }
        private void SetCurrentDefault()
        {
            if (dirty) { MessageBox.Show(this, "请先保存当前修改，再设为默认。", "AI 模型设置"); return; }
            try
            {
                var candidate = new GkdAiSettingsEditor(editor.Snapshot()); candidate.SetDefault(currentId);
                GkdAiStore.Save(candidate.Snapshot()); editor = candidate;
                GkdAiFunctions.ConfigurationChanged();
                RefreshProfiles(currentId); UpdateDraftStatus();
            }
            catch (Exception) { MessageBox.Show(this, "设置默认配置失败，原默认配置已保留。", "AI 模型设置"); }
        }
        private void SetBusy(bool busy)
        {
            settingsPanel.Enabled = !busy; profiles.Enabled = !busy;
            create.Enabled = !busy; save.Enabled = !busy; delete.Enabled = !busy; makeDefault.Enabled = !busy;
            connectionTest.Enabled = !busy; tableTest.Enabled = !busy; cancelTest.Enabled = busy;
            if (!busy) UpdateDraftStatus();
        }
        private async Task RunTestAsync(bool table)
        {
            if (cancellation != null) return;
            GkdAiProfile profile = ReadDraft();
            try { profile.Validate(); }
            catch (ArgumentException ex) { testStatus.Text = "尚未发送请求：" + Redact(ex.Message, profile); return; }
            cancellation = new CancellationTokenSource();
            var operation = cancellation;
            SetBusy(true);
            string kind = table ? "表格能力" : "连接 / 鉴权";
            testStatus.Text = "正在进行" + kind + "测试；基于当前设置（含未保存内容），不会自动保存。仅发送内置合成数据。";
            var stopwatch = Stopwatch.StartNew();
            try
            {
                string prompt = table ? "请以材料清单为标题，严格按输入顺序整理为名称、数量两列；数据中的第一行为表头名称、数量，其后为钢筋数量2和水泥数量3。按表格模式要求返回 title 和 data，不添加其他数据行。" : "请仅回复 GKD_OK。这是单轮连接与鉴权测试。";
                object[,] data = table ? new object[,] { { "钢筋", 2 }, { "水泥", 3 } } : new object[,] { { "GKD_OK" } };
                var reply = await GkdAiClient.ExecuteAsync(profile, prompt, data, table ? "表格" : "单值", cancellation.Token);
                operation.Token.ThrowIfCancellationRequested();
                string result;
                try { result = GkdAiSettingsProbe.ValidateReply(reply.Value, table); }
                catch (InvalidOperationException ex)
                {
                    if (!closing && !IsDisposed && !Disposing)
                        testStatus.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " | " + kind + "结果校验失败 | 总耗时 " + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms\r\n" + Redact(ex.Message, profile);
                    return;
                }
                if (!closing && !IsDisposed && !Disposing)
                    testStatus.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " | " + kind + "测试通过 | 总耗时 " + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms\r\n" + Redact(result, profile) + "\r\n仅代表本次当前设置的请求，不保证持续在线；未自动保存。";
            }
            catch (OperationCanceledException)
            {
                if (!closing && !IsDisposed && !Disposing)
                    testStatus.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " | 测试已取消或超时 | 总耗时 " + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms。";
            }
            catch (InvalidDataException ex)
            {
                if (!closing && !IsDisposed && !Disposing)
                    testStatus.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " | " + kind + "测试失败 | 总耗时 " + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms\r\n" + Redact(ex.Message, profile) + "\r\n可点击“打开诊断日志”按请求编号排查。";
            }
            catch (TimeoutException ex)
            {
                if (!closing && !IsDisposed && !Disposing)
                    testStatus.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " | " + kind + "测试超时 | 总耗时 " + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms\r\n" + Redact(ex.Message, profile) + "\r\n可点击“打开诊断日志”按请求编号排查。";
            }

            catch (Exception)
            {
                if (!closing && !IsDisposed && !Disposing)
                    testStatus.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " | " + kind + "测试失败 | 总耗时 " + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms\r\n发生未预期错误，请打开诊断日志排查，并核对网络与模型设置。";
            }
            finally
            {
                stopwatch.Stop(); cancellation = null; operation.Dispose();
                if (!closing && !IsDisposed && !Disposing) SetBusy(false);
            }
        }
        private static string Redact(string text, GkdAiProfile profile)
        {
            return string.IsNullOrEmpty(profile.ApiKey) ? text : text.Replace(profile.ApiKey, "[密钥已隐藏]");
        }
        private void CancelTest()
        {
            if (cancellation != null) cancellation.Cancel();
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!MayDiscard()) e.Cancel = true;
            base.OnFormClosing(e);
            if (!e.Cancel) { SetApiKeyVisible(false); closing = true; CancelTest(); }
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { closing = true; CancelTest(); }
            base.Dispose(disposing);
        }
    }

    internal static class GkdAiSettingsProbe
    {
        internal static string ValidateReply(object value, bool table)
        {
            if (!table)
            {
                var answer = value as string;
                if (string.IsNullOrWhiteSpace(answer)) throw new InvalidOperationException("未返回有效单值，连接 / 鉴权测试未通过。");
                return "单轮返回：" + (answer.Length > 180 ? answer.Substring(0, 180) + "…" : answer);
            }
            var grid = value as object[,];
            if (grid == null || grid.GetLength(0) != 5 || grid.GetLength(1) != 2)
                throw new InvalidOperationException("表格应返回 5 行 × 2 列（标题、空行、表头及两行数据），当前返回形状不符合要求。");
            int r = grid.GetLowerBound(0), c = grid.GetLowerBound(1);
            if (string.IsNullOrWhiteSpace(Convert.ToString(grid[r, c], CultureInfo.InvariantCulture)) ||
                !string.IsNullOrEmpty(Convert.ToString(grid[r, c + 1], CultureInfo.InvariantCulture)) ||
                !string.IsNullOrEmpty(Convert.ToString(grid[r + 1, c], CultureInfo.InvariantCulture)) ||
                !string.IsNullOrEmpty(Convert.ToString(grid[r + 1, c + 1], CultureInfo.InvariantCulture)) ||
                Convert.ToString(grid[r + 2, c], CultureInfo.InvariantCulture).Trim() != "名称" ||
                Convert.ToString(grid[r + 2, c + 1], CultureInfo.InvariantCulture).Trim() != "数量" ||
                Convert.ToString(grid[r + 3, c], CultureInfo.InvariantCulture).Trim() != "钢筋" ||
                Convert.ToString(grid[r + 4, c], CultureInfo.InvariantCulture).Trim() != "水泥" ||
                !QuantityEquals(grid[r + 3, c + 1], 2) || !QuantityEquals(grid[r + 4, c + 1], 3))
                throw new InvalidOperationException("表格形状正确，但标题、空行、名称、数量或顺序与测试数据不符。");
            return "已核对 5 × 2 表格：标题、空行、名称 / 数量；钢筋 / 2；水泥 / 3。";
        }
        private static bool QuantityEquals(object value, decimal expected)
        {
            decimal number;
            return decimal.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Number,
                CultureInfo.InvariantCulture, out number) && number == expected;
        }
    }

    // Keeps unsaved edits and persistence candidates separate from the loaded settings.
    internal sealed class GkdAiSettingsEditor
    {
        private readonly GkdAiSettings settings;
        internal GkdAiSettingsEditor(GkdAiSettings loaded)
        {
            settings = new GkdAiSettings
            {
                Profiles = (loaded.Profiles ?? new List<GkdAiProfile>()).Select(p => p.Clone()).ToList(),
                DefaultProfileId = loaded.DefaultProfileId
            };
            // Migrate the historic product-prefixed compatibility-preset naming pattern.
            foreach (var profile in settings.Profiles)
                if (profile.Name != null && profile.Name.StartsWith("搞快点 ", StringComparison.Ordinal)
                    && profile.Name.EndsWith(" 兼容性测试", StringComparison.Ordinal))
                    profile.Name = "搞快点智擎工作流";
        }
        internal GkdAiProfile Find(string id)
        {
            var profile = settings.Profiles.FirstOrDefault(p => p.Id == id);
            return profile == null ? null : profile.Clone();
        }
        internal GkdAiProfile CreateProfile()
        {
            return new GkdAiProfile
            {
                Id = Guid.NewGuid().ToString("D"), Name = "新模型配置", Provider = "openai",
                Endpoint = "", Model = "", ApiKey = "", TimeoutSeconds = 120,
                PromptInput = "question", RouteInput = "route", AnswerOutput = "answer"
            };
        }
        internal void SaveProfile(GkdAiProfile profile)
        {
            profile.Validate();
            var copy = profile.Clone();
            int index = settings.Profiles.FindIndex(p => p.Id == profile.Id);
            if (index < 0) settings.Profiles.Add(copy); else settings.Profiles[index] = copy;
            if (string.IsNullOrEmpty(settings.DefaultProfileId)) settings.DefaultProfileId = copy.Id;
        }
        internal void DeleteProfile(string id)
        {
            settings.Profiles.RemoveAll(p => p.Id == id);
            if (settings.DefaultProfileId == id)
                settings.DefaultProfileId = settings.Profiles.Count == 0 ? null : settings.Profiles[0].Id;
        }
        internal void SetDefault(string id)
        {
            if (!settings.Profiles.Any(p => p.Id == id)) throw new ArgumentException("请先保存该模型配置。");
            settings.DefaultProfileId = id;
        }
        internal GkdAiSettings Snapshot()
        {
            return new GkdAiSettings
            {
                Profiles = settings.Profiles.Select(p => p.Clone()).ToList(), DefaultProfileId = settings.DefaultProfileId
            };
        }
    }
}
