using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Outline.Addin
{
    // Declarative definitions only; numeric trials use the restricted evaluator, never source code.
    internal sealed class GkdCustomFormulaForm : Form
    {
        private readonly ListBox formulas = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, DisplayMember = "DisplayName", FormattingEnabled = true };
        private readonly TextBox details = TextArea(true);
        private readonly TextBox status = TextArea(true);
        private readonly Button create = Ui.Button("新建");
        private readonly Button edit = Ui.Button("编辑");
        private readonly Button delete = Ui.Button("删除");
        private readonly Button import = Ui.Button("导入");
        private readonly Button export = Ui.Button("导出");
        private List<GkdCustomFormula> saved;
        private bool loadFailed;
        private GkdCustomFormulaEditorForm activeEditor;

        internal GkdCustomFormulaForm()
        {
            Ui.Base(this, new Size(1060, 720), "自定义公式 — 本地公式管理");
            MinimumSize = new Size(900, 620);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            root.Controls.Add(new Label { Text = "在本机管理公式的含义、参数和算法定义。AI 设计结果必须先预览并由你确认，才会保存。\r\n此窗口不运行生成的代码，也不修改工作簿。", Dock = DockStyle.Fill, AutoSize = true }, 0, 0);
            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 67));
            body.Controls.Add(formulas, 0, 0);
            body.Controls.Add(details, 1, 0);
            root.Controls.Add(body, 0, 1);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
            actions.Controls.AddRange(new Control[] { create, edit, delete, import, export });
            root.Controls.Add(actions, 0, 2);
            var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footer.Controls.Add(status, 0, 0);
            var close = Ui.Button("关闭"); close.Click += (s, e) => Close();
            footer.Controls.Add(close, 1, 0);
            root.Controls.Add(footer, 0, 3);
            Controls.Add(root);
            CancelButton = close;
            formulas.Format += (s, e) => { var item = e.ListItem as GkdCustomFormula; if (item != null) e.Value = item.DisplayName + "（" + item.ExcelName + "）"; };
            formulas.SelectedIndexChanged += (s, e) => ShowSelection();
            formulas.DoubleClick += (s, e) => EditCurrent();
            create.Click += (s, e) => OpenEditor(null);
            edit.Click += (s, e) => EditCurrent();
            delete.Click += (s, e) => DeleteCurrent();
            import.Click += (s, e) => ImportDefinitions();
            export.Click += (s, e) => ExportDefinitions();
            try { saved = GkdCustomFormulaStore.Load(); }
            catch (Exception) { saved = new List<GkdCustomFormula>(); loadFailed = true; }
            RefreshList(null);
            if (loadFailed) status.Text = "公式库读取失败。原文件未修改；请备份或修复后重新打开。";
        }

        private void RefreshList(string id)
        {
            formulas.Items.Clear();
            foreach (var item in saved) formulas.Items.Add(Copy(item));
            if (id != null)
                for (int i = 0; i < formulas.Items.Count; i++)
                    if (((GkdCustomFormula)formulas.Items[i]).Id == id) formulas.SelectedIndex = i;
            if (formulas.SelectedIndex < 0 && formulas.Items.Count > 0) formulas.SelectedIndex = 0;
            create.Enabled = import.Enabled = !loadFailed;
            ShowSelection();
            status.Text = saved.Count == 0 ? "暂无自定义公式。点击“新建”，填写算法或使用 AI 设计。" : "本地已保存 " + saved.Count + " 个公式。";
        }

        private void ShowSelection()
        {
            var item = formulas.SelectedItem as GkdCustomFormula;
            edit.Enabled = delete.Enabled = item != null && !loadFailed;
            export.Enabled = saved.Count > 0 && !loadFailed;
            details.Text = item == null ? "选择一个公式查看含义、参数和算法定义。" : Describe(item);
        }

        private void EditCurrent()
        {
            var selected = formulas.SelectedItem as GkdCustomFormula;
            if (selected != null && !loadFailed) OpenEditor(selected);
        }

        private void OpenEditor(GkdCustomFormula original)
        {
            if (loadFailed) return;
            using (var editor = new GkdCustomFormulaEditorForm(original, candidate => SaveDraft(candidate)))
            {
                activeEditor = editor;
                try { editor.ShowDialog(this); }
                finally { activeEditor = null; }
            }
        }

        internal void CloseForShutdown()
        {
            if (activeEditor != null && !activeEditor.IsDisposed) activeEditor.CloseForShutdown();
            Close();
        }

        private void SaveDraft(GkdCustomFormula draft)
        {
            ValidateForReview(draft);
            var candidate = saved.Select(Copy).ToList();
            int index = candidate.FindIndex(x => x.Id == draft.Id);
            if (index < 0) candidate.Add(GkdCustomFormulaForm.Copy(draft)); else candidate[index] = GkdCustomFormulaForm.Copy(draft);
            GkdCustomFormulaStore.Save(candidate);
            GkdCustomFormulaFunctions.Reload();
            saved = candidate;
            RefreshList(draft.Id);
        }

        private void DeleteCurrent()
        {
            var selected = formulas.SelectedItem as GkdCustomFormula;
            if (selected == null || loadFailed) return;
            if (MessageBox.Show(this, "删除本地公式“" + selected.DisplayName + "”？\r\n" + selected.Description + "\r\n已使用此定义的公式可能受影响。", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            try
            {
                var candidate = saved.Where(x => x.Id != selected.Id).Select(Copy).ToList();
                GkdCustomFormulaStore.Save(candidate);
                try { GkdCustomFormulaFunctions.Reload(); }
                catch (Exception) { status.Text = "公式已保存，但本次会话未能重新注册；重启 Excel 后重试。"; }
                saved = candidate;
                RefreshList(null);
            }
            catch (Exception) { ShowError("删除失败，原公式库已保留。请检查本地目录的写入权限。"); }
        }

        private void ImportDefinitions()
        {
            if (loadFailed) return;
            using (var picker = new OpenFileDialog { Filter = "公式定义 JSON (*.json)|*.json", CheckFileExists = true, Multiselect = false })
            {
                if (picker.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    if (new FileInfo(picker.FileName).Length > 1024 * 1024) throw new InvalidDataException("导入文件不能超过 1 MB。");
                    var incoming = GkdCustomFormulaJson.Import(File.ReadAllText(picker.FileName));
                    if (incoming.Count == 0) throw new InvalidDataException("文件中没有公式。");
                    var candidate = saved.Select(Copy).ToList();
                    foreach (var item in incoming)
                    {
                        ValidateForReview(item);
                        if (candidate.Any(x => x.Id == item.Id || string.Equals(x.ExcelName, item.ExcelName, StringComparison.OrdinalIgnoreCase)))
                            throw new InvalidDataException("导入项与现有公式标识或名称重复。请先编辑名称或删除旧项，再重新导入。");
                        candidate.Add(Copy(item));
                    }
                    using (var preview = new GkdCustomFormulaPreviewForm("导入预览 — " + incoming.Count + " 个公式", string.Join("\r\n\r\n────────\r\n", incoming.Select(Describe)), true))
                    {
                        if (preview.ShowDialog(this) != DialogResult.OK) return;
                        GkdCustomFormulaStore.Save(candidate);
                        try { GkdCustomFormulaFunctions.Reload(); }
                        catch (Exception) { status.Text = "公式已导入并保存，但本次会话未能重新注册；重启 Excel 后重试。"; }
                        saved = candidate;
                        RefreshList(incoming[0].Id);
                    }
                }
                catch (ArgumentException ex) { ShowError(ex.Message); }
                catch (InvalidDataException ex) { ShowError(ex.Message); }
                catch (Exception) { ShowError("导入失败，原公式库未修改。请检查文件格式及访问权限。"); }
            }
        }

        private void ExportDefinitions()
        {
            using (var picker = new SaveFileDialog { Filter = "公式定义 JSON (*.json)|*.json", FileName = "自定义公式.json", DefaultExt = "json", AddExtension = true, OverwritePrompt = true })
            {
                if (picker.ShowDialog(this) != DialogResult.OK) return;
                try { File.WriteAllText(picker.FileName, GkdCustomFormulaJson.Export(saved.Select(Copy).ToList())); status.Text = "已导出全部 " + saved.Count + " 个公式。"; }
                catch (Exception) { ShowError("导出失败，请检查目标文件的写入权限。"); }
            }
        }

        private void ShowError(string message) { MessageBox.Show(this, message, "自定义公式", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        internal static string Describe(GkdCustomFormula item)
        {
            string arguments = string.Join(",", Enumerable.Range(0, item.Inputs.Count).Select(i => (i < 26 ? ((char)('A' + i)).ToString() : "A" + (char)('A' + i - 26)) + "2"));
            return "名称：" + item.DisplayName + "（" + item.ExcelName + "）\r\n类型：" + (item.Kind == GkdCustomFormulaKind.Ai ? "AI 提示词" : "确定性表达式")
                + "；版本：" + item.Version + "；" + (item.Enabled ? "已启用" : "已停用")
                + "\r\n\r\n含义说明：\r\n" + item.Description + "\r\n\r\n参数：\r\n" + string.Join(", ", item.Inputs)
                + "\r\n输出单位：" + item.OutputUnit + "\r\n\r\nExcel 调用示例（按参数声明顺序）：\r\n=" + item.ExcelName + "(" + arguments + ")\r\n=GKD_CUSTOM(\"" + item.ExcelName + "\"" + (arguments.Length == 0 ? "" : "," + arguments) + ")"
                + "\r\n\r\n算法定义：\r\n" + (item.Kind == GkdCustomFormulaKind.Ai ? item.PromptTemplate : item.FormulaExpression);
        }
        internal static GkdCustomFormula Copy(GkdCustomFormula item)
        {
            return new GkdCustomFormula { Id = item.Id, ExcelName = item.ExcelName, DisplayName = item.DisplayName, Description = item.Description,
                Kind = item.Kind, Inputs = new List<string>(item.Inputs), FormulaExpression = item.FormulaExpression,
                PromptTemplate = item.PromptTemplate, OutputUnit = item.OutputUnit, Version = item.Version, Enabled = item.Enabled };
        }
        internal static void ValidateForReview(GkdCustomFormula item)
        {
            item.Validate();
            if (string.IsNullOrWhiteSpace(item.Description)) throw new ArgumentException("请填写含义说明，并解释参数单位、适用范围及边界条件。");
        }
        internal static TextBox TextArea(bool readOnly)
        {
            return new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = readOnly, ScrollBars = ScrollBars.Vertical, BackColor = Color.White, AcceptsReturn = true, MaxLength = 65536 };
        }
    }

    internal sealed class GkdCustomFormulaEditorForm : Form
    {
        private readonly TextBox name = new TextBox { Dock = DockStyle.Fill, MaxLength = 100 };
        private readonly TextBox excelName = new TextBox { Dock = DockStyle.Fill, MaxLength = 31 };
        private readonly TextBox unit = new TextBox { Dock = DockStyle.Fill, MaxLength = 32 };
        private readonly ComboBox kind = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly CheckBox enabled = new CheckBox { Text = "启用此定义", AutoSize = true, Checked = true };
        private readonly TextBox meaning = GkdCustomFormulaForm.TextArea(false);
        private readonly TextBox parameters = GkdCustomFormulaForm.TextArea(false);
        private readonly TextBox definition = GkdCustomFormulaForm.TextArea(false);
        private readonly TextBox request = GkdCustomFormulaForm.TextArea(false);
        private readonly TextBox samples = new TextBox { Dock = DockStyle.Fill, MaxLength = 4096 };
        private readonly TextBox status = GkdCustomFormulaForm.TextArea(true);
        private readonly Button design = Ui.Button("AI 设计");
        private readonly Button trial = Ui.Button("试算");
        private readonly Button save = Ui.Button("预览并确认保存", true);
        private readonly Button cancel = Ui.Button("取消 AI 请求");
        private readonly Action<GkdCustomFormula> persist;
        private readonly string id;
        private readonly int version;
        private readonly Control fields;
        private CancellationTokenSource cancellation;
        private bool dirty;
        private bool closing;
        private bool shuttingDown;

        internal GkdCustomFormulaEditorForm(GkdCustomFormula original, Action<GkdCustomFormula> persist)
        {
            this.persist = persist;
            var draft = original == null ? new GkdCustomFormula() : GkdCustomFormulaForm.Copy(original);
            id = draft.Id;
            version = draft.Version;
            Ui.Base(this, new Size(1020, 900), original == null ? "新建自定义公式" : "编辑自定义公式");
            MinimumSize = new Size(950, 820);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 96));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
            var inputs = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 10 };
            inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 118));
            inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            inputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            inputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            inputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            inputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            inputs.RowStyles.Add(new RowStyle(SizeType.Percent, 18));
            inputs.RowStyles.Add(new RowStyle(SizeType.Percent, 18));
            inputs.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            inputs.RowStyles.Add(new RowStyle(SizeType.Percent, 24));
            inputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            inputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
            AddField(inputs, 0, "公式名称", name);
            AddField(inputs, 1, "Excel 标识", excelName);
            kind.Items.AddRange(new object[] { "确定性表达式", "AI 提示词" });
            AddField(inputs, 2, "定义类型", kind);
            var output = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            unit.Dock = DockStyle.None; unit.Width = 180; output.Controls.AddRange(new Control[] { unit, enabled });
            AddField(inputs, 3, "输出单位", output);
            AddField(inputs, 4, "含义 / 参数单位", meaning);
            AddField(inputs, 5, "输入名（逗号分隔）", parameters);
            AddField(inputs, 6, "表达式 / 提示词", definition);
            AddField(inputs, 7, "AI 设计需求", request);
            AddField(inputs, 8, "样例值（数值）", samples);
            inputs.Controls.Add(new Label { Text = "标识 / 输入名使用 A-Z、数字、下划线并以字母开头。受限表达式支持数字、输入名、+ - * / ^、括号和 CONVERT(值,\"mm\",\"m\")；仅同维度的受支持单位可换算，不支持其他函数或代码。\r\nPI 不是内置常量，请填写数字 3.141592653589793。请解释参数单位及适用范围。\r\n样例值按输入名声明顺序，以逗号分隔；小数用点号，例如 120,8；零输入留空。AI 试算将预览并发送当前提示词和带参数名的数值。\r\n" + GkdCustomFormulaTrial.Limitation, AutoSize = true, Dock = DockStyle.Fill }, 0, 9);
            inputs.SetColumnSpan(inputs.GetControlFromPosition(0, 9), 2);
            fields = inputs;
            root.Controls.Add(inputs, 0, 0);
            root.Controls.Add(status, 0, 1);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = true };
            var close = Ui.Button("关闭"); close.Click += (s, e) => Close();
            actions.Controls.AddRange(new Control[] { design, trial, cancel, save, close });
            root.Controls.Add(actions, 0, 2);
            Controls.Add(root);
            CancelButton = close;
            SetDraftFields(draft);
            foreach (var field in new[] { name, excelName, unit, meaning, parameters, definition, request }) field.TextChanged += (s, e) => MarkDirty();
            kind.SelectedIndexChanged += (s, e) => MarkDirty();
            enabled.CheckedChanged += (s, e) => MarkDirty();
            design.Click += async (s, e) => await DesignAsync();
            trial.Click += async (s, e) => await TrialAsync();
            cancel.Click += (s, e) => { if (cancellation != null) cancellation.Cancel(); };
            save.Click += (s, e) => SaveCurrent();
            SetBusy(false);
            status.Text = "请核对含义、参数、单位和算法定义。";
        }

        private static void AddField(TableLayoutPanel table, int row, string label, Control field)
        {
            table.Controls.Add(new Label { Text = label, Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 6, 8, 4) }, 0, row);
            table.Controls.Add(field, 1, row);
        }
        private GkdCustomFormula ReadDraft()
        {
            return new GkdCustomFormula { Id = id, ExcelName = excelName.Text.Trim(), DisplayName = name.Text.Trim(), Description = meaning.Text.Trim(),
                Inputs = parameters.Text.Split(new[] { ',', '，', ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList(),
                OutputUnit = unit.Text.Trim(), Kind = kind.SelectedIndex == 1 ? GkdCustomFormulaKind.Ai : GkdCustomFormulaKind.Deterministic,
                FormulaExpression = kind.SelectedIndex == 1 ? "" : definition.Text.Trim(), PromptTemplate = kind.SelectedIndex == 1 ? definition.Text.Trim() : "",
                Enabled = enabled.Checked, Version = version };
        }
        private void MarkDirty() { dirty = true; status.Text = "草稿尚未保存；修改后需要重新预览确认。"; }
        private void SetDraftFields(GkdCustomFormula draft)
        {
            name.Text = draft.DisplayName; excelName.Text = draft.ExcelName; meaning.Text = draft.Description;
            parameters.Text = string.Join(", ", draft.Inputs); unit.Text = draft.OutputUnit; enabled.Checked = draft.Enabled;
            kind.SelectedIndex = draft.Kind == GkdCustomFormulaKind.Ai ? 1 : 0;
            definition.Text = draft.Kind == GkdCustomFormulaKind.Ai ? draft.PromptTemplate : draft.FormulaExpression;
        }
        private void SaveCurrent()
        {
            if (cancellation != null) return;
            try
            {
                var draft = ReadDraft(); GkdCustomFormulaForm.ValidateForReview(draft);
                using (var preview = new GkdCustomFormulaPreviewForm("确认保存自定义公式", GkdCustomFormulaForm.Describe(draft), true))
                {
                    if (preview.ShowDialog(this) != DialogResult.OK) return;
                    persist(GkdCustomFormulaForm.Copy(draft));
                }
                dirty = false; DialogResult = DialogResult.OK; Close();
            }
            catch (ArgumentException ex) { status.Text = "未保存：" + ex.Message; }
            catch (InvalidDataException ex) { status.Text = "未保存：" + ex.Message; }
            catch (Exception) { status.Text = "保存失败，草稿仍保留在窗口中。请检查本地公式库和写入权限。"; }
        }

        private async Task TrialAsync()
        {
            if (cancellation != null) return;
            GkdCustomFormula draft;
            GkdAiProfile profile;
            object[,] data;
            try
            {
                draft = ReadDraft();
                if (draft.Kind == GkdCustomFormulaKind.Deterministic)
                {
                    status.Text = GkdCustomFormulaTrial.Evaluate(draft, samples.Text);
                    return;
                }
                var values = GkdCustomFormulaTrial.ParseValues(draft, samples.Text);
                if (!draft.Enabled) throw new ArgumentException("公式未启用；请先启用再试算。");
                data = GkdCustomFormulaTrial.LabelledData(draft, values);
                var settings = GkdAiStore.Load();
                profile = GkdAiStore.ResolveDefault(settings);
                profile = profile.Clone(); profile.Validate();
            }
            catch (ArgumentException ex) { status.Text = "无法试算：" + ex.Message; return; }
            catch (Exception) { status.Text = "无法试算：读取默认 AI 配置失败，请检查 AI 设置。"; return; }
            string previewText = GkdAiProtocol.BuildPrompt(draft.PromptTemplate, data, "单值");
            using (var preview = new GkdCustomFormulaPreviewForm("AI 试算提示词与数据预览 — " + profile.Name, previewText, false))
                if (preview.ShowDialog(this) != DialogResult.OK) return;
            var tokenSource = new CancellationTokenSource();
            cancellation = tokenSource; SetBusy(true); status.Text = "正在进行 AI 样例试算，可取消；不会保存或修改定义。";
            try
            {
                var reply = await GkdAiClient.ExecuteAsync(profile, draft.PromptTemplate, data, "单值", tokenSource.Token);
                tokenSource.Token.ThrowIfCancellationRequested();
                if (closing || IsDisposed || Disposing) return;
                string answer = reply.Value as string;
                if (string.IsNullOrWhiteSpace(answer)) throw new InvalidDataException("AI 未返回单值试算结果。");
                status.Text = "AI 试算结果：" + answer + "；单位：" + (string.IsNullOrWhiteSpace(draft.OutputUnit) ? "未声明" : draft.OutputUnit) + "\r\n" + GkdCustomFormulaTrial.Limitation;
            }
            catch (OperationCanceledException) { if (!closing && !IsDisposed) status.Text = "已取消 AI 试算；当前草稿和公式库未改变。"; }
            catch (InvalidDataException ex) { if (!closing && !IsDisposed) status.Text = "AI 试算失败：" + ex.Message; }
            catch (TimeoutException ex) { if (!closing && !IsDisposed) status.Text = "AI 试算超时：" + ex.Message; }
            catch (Exception) { if (!closing && !IsDisposed) status.Text = "AI 试算未完成，草稿已保留。请检查默认模型配置后重试。"; }
            finally
            {
                if (ReferenceEquals(cancellation, tokenSource)) cancellation = null;
                tokenSource.Dispose();
                if (!closing && !IsDisposed && !Disposing) SetBusy(false);
            }
        }

        private async Task DesignAsync()
        {
            if (cancellation != null) return;
            if (string.IsNullOrWhiteSpace(request.Text)) { status.Text = "请填写 AI 设计需求，包括参数、单位及计算规则。"; return; }
            GkdAiProfile profile;
            string prompt;
            try
            {
                var settings = GkdAiStore.Load();
                profile = GkdAiStore.ResolveDefault(settings);
                profile = profile.Clone(); profile.Validate();
                prompt = GkdCustomFormulaDesignPrompt.Build(ReadDraft(), request.Text);
            }
            catch (ArgumentException ex) { status.Text = "尚未发送请求：" + ex.Message; return; }
            catch (Exception) { status.Text = "尚未发送请求。请检查 AI 设置中的默认模型、接口和密钥。"; return; }
            using (var preview = new GkdCustomFormulaPreviewForm("AI 设计提示词预览 — " + profile.Name, prompt, false))
                if (preview.ShowDialog(this) != DialogResult.OK) return;
            var tokenSource = new CancellationTokenSource();
            cancellation = tokenSource; SetBusy(true); status.Text = "正在生成算法草案，可取消；生成结果不会自动保存。";
            try
            {
                var reply = await GkdAiClient.ExecuteAsync(profile, prompt, null, "单值", tokenSource.Token);
                tokenSource.Token.ThrowIfCancellationRequested();
                if (closing || IsDisposed || Disposing) return;
                var text = reply.Value as string;
                if (text == null) throw new InvalidDataException("AI 未返回单值算法定义。");
                var drafts = GkdCustomFormulaJson.Import(text);
                if (drafts.Count != 1) throw new InvalidDataException("AI 必须只返回一个公式草案。");
                var draft = drafts[0]; GkdCustomFormulaForm.ValidateForReview(draft);
                SetDraftFields(draft);
                dirty = true;
                status.Text = "AI 草案已填入，尚未保存。请逐项核对含义、参数和算法，再点击“预览并确认保存”。";
            }
            catch (OperationCanceledException) { if (!closing && !IsDisposed) status.Text = "已取消 AI 设计；当前草稿及本地公式库未改变。"; }
            catch (InvalidDataException ex) { if (!closing && !IsDisposed) status.Text = "AI 设计失败：" + ex.Message; }
            catch (TimeoutException ex) { if (!closing && !IsDisposed) status.Text = "AI 设计超时：" + ex.Message; }
            catch (Exception) { if (!closing && !IsDisposed) status.Text = "AI 设计失败或定义格式不受支持。当前草稿和本地公式库已保留，请检查设置或重试。"; }
            finally
            {
                if (ReferenceEquals(cancellation, tokenSource)) cancellation = null;
                tokenSource.Dispose();
                if (!closing && !IsDisposed && !Disposing) SetBusy(false);
            }
        }
        private void SetBusy(bool busy) { fields.Enabled = design.Enabled = trial.Enabled = save.Enabled = !busy; cancel.Enabled = busy; }
        internal void CloseForShutdown() { shuttingDown = true; closing = true; if (cancellation != null) cancellation.Cancel(); Close(); }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!shuttingDown && e.CloseReason != CloseReason.ApplicationExitCall && dirty && DialogResult != DialogResult.OK && MessageBox.Show(this, "当前草稿尚未保存，是否放弃修改？", "自定义公式", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) e.Cancel = true;
            if (!e.Cancel) { closing = true; if (cancellation != null) cancellation.Cancel(); }
            base.OnFormClosing(e);
        }
        protected override void Dispose(bool disposing) { if (disposing) { closing = true; if (cancellation != null) cancellation.Cancel(); } base.Dispose(disposing); }
    }

    internal sealed class GkdCustomFormulaPreviewForm : Form
    {
        internal GkdCustomFormulaPreviewForm(string title, string content, bool saving)
        {
            Ui.Base(this, new Size(880, 680), title);
            MinimumSize = new Size(720, 530);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            var preview = GkdCustomFormulaForm.TextArea(true); preview.Text = content;
            root.Controls.Add(preview, 0, 0);
            var reviewed = new CheckBox { Text = saving ? "我已核对所有公式的含义、参数、单位和算法定义，确认保存。" : "我已核对提示词，允许发送到当前默认模型。", AutoSize = true, Dock = DockStyle.Fill, Checked = false };
            root.Controls.Add(reviewed, 0, 1);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            var cancel = Ui.Button("取消"); cancel.DialogResult = DialogResult.Cancel;
            var confirm = Ui.Button(saving ? "确认保存" : "发送 AI 请求", true); confirm.Enabled = false;
            reviewed.CheckedChanged += (s, e) => confirm.Enabled = reviewed.Checked;
            confirm.Click += (s, e) => { if (!reviewed.Checked) return; DialogResult = DialogResult.OK; Close(); };
            actions.Controls.AddRange(new Control[] { cancel, confirm });
            root.Controls.Add(actions, 0, 2);
            Controls.Add(root); CancelButton = cancel;
        }
    }

    internal static class GkdCustomFormulaTrial
    {
        internal const string Limitation = "样例试算不能证明算法正确；请另行核对单位、适用范围和边界条件。";

        internal static Dictionary<string, double> ParseValues(GkdCustomFormula draft, string samples)
        {
            draft.Validate();
            string[] parts = string.IsNullOrWhiteSpace(samples) ? new string[0] : samples.Split(new[] { ',', '，', ';', '；', '\r', '\n' }, StringSplitOptions.None);
            if (parts.Length != draft.Inputs.Count) throw new ArgumentException("样例值数量必须与输入名一致（" + draft.Inputs.Count + " 个），按声明顺序逐项填写；零输入请留空。");
            var values = new Dictionary<string, double>(StringComparer.Ordinal);
            for (int i = 0; i < parts.Length; i++)
            {
                double value;
                if (!double.TryParse(parts[i].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value) || double.IsNaN(value) || double.IsInfinity(value))
                    throw new ArgumentException("参数 " + draft.Inputs[i] + " 的样例值必须为有限数字；小数使用点号，数值之间使用逗号。");
                values.Add(draft.Inputs[i], value);
            }
            return values;
        }

        internal static string Evaluate(GkdCustomFormula draft, string samples)
        {
            double result = draft.Evaluate(ParseValues(draft, samples));
            if (double.IsNaN(result) || double.IsInfinity(result)) throw new ArgumentException("试算结果不是有限数字。");
            return "试算结果：" + result.ToString("G17", CultureInfo.InvariantCulture) + "；单位：" + (string.IsNullOrWhiteSpace(draft.OutputUnit) ? "未声明" : draft.OutputUnit) + "\r\n" + Limitation;
        }

        internal static object[,] LabelledData(GkdCustomFormula draft, IDictionary<string, double> values)
        {
            if (draft.Inputs.Count == 0) return null;
            var data = new object[draft.Inputs.Count, 2];
            for (int i = 0; i < draft.Inputs.Count; i++) { data[i, 0] = draft.Inputs[i]; data[i, 1] = values[draft.Inputs[i]]; }
            return data;
        }
    }

    internal static class GkdCustomFormulaDesignPrompt
    {
        internal static string Build(GkdCustomFormula draft, string request)
        {
            if (string.IsNullOrWhiteSpace(request) || request.Length > 8000) throw new ArgumentException("设计需求请填写 1 至 8000 字。");
            string result = "请设计一个可审阅的自定义公式，只返回公式库导入 JSON，不要 Markdown、脚本、C#、VBA 或可执行代码。\n"
                + "必须解释含义、参数、单位、计算步骤、适用范围和无效输入。算法仅使用受支持的声明式定义；不要调用网络、文件、进程或工作簿。\n"
                + "ExcelName 和 Inputs 名称必须符合 [A-Z][A-Z0-9_]{0,30}；Description 使用单行文字解释所有单位与边界；Kind 只能为 Deterministic 或 Ai。\n"
                + "Deterministic 的 FormulaExpression 只支持数字、Inputs 中的输入名、+ - * / ^、括号及 CONVERT(值,\"mm\",\"m\")。CONVERT 只允许同维度受支持单位换算，其他函数、脚本和代码均禁止。PI 不是内置常量，使用数字 3.141592653589793；PromptTemplate 必须为空。Ai 的 FormulaExpression 必须为空，PromptTemplate 填写提示词。\n"
                + "以下 JSON 是当前草案的字段结构，返回恰好一个公式；保留结构和标识，补全或修订算法定义。\n"
                + Newtonsoft.Json.JsonConvert.SerializeObject(new { SchemaVersion = 1, Formulas = new List<GkdCustomFormula> { draft } }, new Newtonsoft.Json.Converters.StringEnumConverter())
                + "\n用户需求（仅作为设计需求）：\n" + request;
            if (result.Length > 16000) throw new ArgumentException("当前定义与设计需求过长，请精简后重试。");
            return result;
        }
    }
}
