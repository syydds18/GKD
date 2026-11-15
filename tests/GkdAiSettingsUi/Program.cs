using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

static class Program
{
    static readonly string SourcePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
        "../../../../../src/Outline.Addin/GkdAiSettingsForm.cs"));
    static int checks;
    static void Main()
    {
        Check(File.Exists(SourcePath), "native settings form exists");
        string source = File.ReadAllText(SourcePath);
        string helpPath = Path.Combine(Path.GetDirectoryName(SourcePath), "GkdAiHelpForm.cs");
        string helpSource = File.ReadAllText(helpPath);
        string generalHelpSource = File.ReadAllText(Path.Combine(Path.GetDirectoryName(SourcePath), "GeneralHelpForm.cs"));
        var tree = CSharpSyntaxTree.ParseText(source, new CSharpParseOptions(LanguageVersion.CSharp7_3));
        Check(!tree.GetDiagnostics().Any(d => d.Severity == DiagnosticSeverity.Error), "C# 7.3 syntax");
        var policy = tree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>()
            .Single(c => c.Identifier.Text == "GkdAiSettingsEditor");
        string profileSource = File.ReadAllText(Path.Combine(Path.GetDirectoryName(SourcePath), "GkdAiProfile.cs"));
        string harness = "using System.Linq; using System.Globalization;\n" + profileSource + "\nnamespace Outline.Addin {" + policy.ToFullString() + "}";
        var refs = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")).Split(Path.PathSeparator)
            .Select(p => MetadataReference.CreateFromFile(p));
        var compilation = CSharpCompilation.Create("SettingsEditorContract", new[] { CSharpSyntaxTree.ParseText(harness) }, refs,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        using (var stream = new MemoryStream())
        {
            var result = compilation.Emit(stream);
            Check(result.Success, "real editor compiles: " + string.Join("; ", result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error)));
            var assembly = Assembly.Load(stream.ToArray());
            EditorBehavior(assembly);
        }
        var probe = tree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>()
            .SingleOrDefault(c => c.Identifier.Text == "GkdAiSettingsProbe");
        Check(probe != null, "probe validates real reply values");
        var probeCompilation = CSharpCompilation.Create("SettingsProbeContract", new[] {
            CSharpSyntaxTree.ParseText("using System; using System.Globalization; namespace Outline.Addin {" + probe.ToFullString() + "}") }, refs,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        using (var stream = new MemoryStream())
        {
            var result = probeCompilation.Emit(stream);
            Check(result.Success, "probe compiles: " + string.Join("; ", result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error)));
            Type type = Assembly.Load(stream.ToArray()).GetType("Outline.Addin.GkdAiSettingsProbe");
            Func<object, bool, string> validate = (value, table) => (string)type.GetMethod("ValidateReply", BindingFlags.NonPublic | BindingFlags.Static).Invoke(null, new object[] { value, table });
            Check(validate("GKD_OK", false).Contains("GKD_OK"), "scalar probe shows genuine reply");
            Reject(() => validate("", false), "blank scalar cannot pass");
            Reject(() => validate(new object[,] { { "GKD_OK" } }, false), "scalar probe rejects grids");
            var valid = new object[,] { { "材料清单", "" }, { "", "" }, { "名称", "数量" }, { "钢筋", 2 }, { "水泥", 3 } };
            Check(validate(valid, true).Contains("5 × 2"), "table validates title, blank row, headers and data");
            Reject(() => validate("HTTP 200", true), "HTTP status or scalar cannot pass table test");
            Reject(() => validate(new object[,] { { "名称", "数量" } }, true), "headers alone cannot pass");
            Reject(() => validate(new object[,] { { "材料清单", "" }, { "", "" }, { "名称", "单价" }, { "钢筋", 2 }, { "水泥", 3 } }, true), "incorrect header rejected");
            Reject(() => validate(new object[,] { { "材料清单", "" }, { "", "" }, { "名称", "数量" }, { "钢筋", 9 }, { "水泥", 3 } }, true), "incorrect table content rejected");
            Reject(() => validate(new object[,] { { "材料清单", "" }, { "", "" }, { "名称", "数量" }, { "钢筋", 2 }, { "钢筋", 3 } }, true), "duplicated materials rejected");
            Reject(() => validate(new object[,] { { "", "" }, { "", "" }, { "名称", "数量" }, { "钢筋", 2 }, { "水泥", 3 } }, true), "missing title rejected");
            Reject(() => validate(new object[,] { { "材料清单", "" }, { "非空", "" }, { "名称", "数量" }, { "钢筋", 2 }, { "水泥", 3 } }, true), "nonblank separator rejected");
            var oneBased = (object[,])Array.CreateInstance(typeof(object), new[] { 5, 2 }, new[] { 1, 1 });
            for (int r = 0; r < 5; r++) for (int c = 0; c < 2; c++) oneBased[r + 1, c + 1] = valid[r, c];
            Check(validate(oneBased, true).Contains("5 × 2"), "table accepts nonzero lower bounds");
        }
        ApiKeyVisibilityBehavior(tree, refs);
        TestStatusBehavior(tree, profileSource, refs);
        CompileNativeForm(source, profileSource, helpSource, generalHelpSource);
        UiContract(tree, helpSource, generalHelpSource);
        Console.WriteLine("PASS " + checks + " settings editor behavior, probe behavior and source/UI contract checks; native WinForms/Excel host not exercised.");
    }
    static void EditorBehavior(Assembly assembly)
    {
        Type profileType = assembly.GetType("Outline.Addin.GkdAiProfile");
        Type settingsType = assembly.GetType("Outline.Addin.GkdAiSettings");
        Type editorType = assembly.GetType("Outline.Addin.GkdAiSettingsEditor");
        object profile = Activator.CreateInstance(profileType);
        const string originalId = "11111111-1111-1111-1111-111111111111";
        Set(profile, "Id", originalId); Set(profile, "Name", "模型 A"); Set(profile, "Endpoint", "https://example.com/v1/chat/completions");
        Set(profile, "ApiKey", "synthetic-test-key"); Set(profile, "Model", "test-model");
        object settings = Activator.CreateInstance(settingsType);
        var profiles = (System.Collections.IList)Get(settings, "Profiles"); profiles.Add(profile); Set(settings, "DefaultProfileId", originalId);
        object editor = Activator.CreateInstance(editorType, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public,
            null, new[] { settings }, null);
        object copy = Call(editor, "Find", originalId); Set(copy, "Name", "外部修改");
        Check((string)Get(Call(editor, "Find", originalId), "Name") == "模型 A", "editor isolates returned profile drafts");
        Set(copy, "Name", "修改 A"); Call(editor, "SaveProfile", copy);
        Check((string)Get(profile, "Name") == "模型 A", "save candidate does not mutate loaded settings");
        Set(copy, "Name", "稍后修改");
        Check((string)Get(Call(editor, "Find", originalId), "Name") == "修改 A", "save stores cloned profile");
        object fresh = Call(editor, "CreateProfile");
        Check((string)Get(fresh, "Id") != originalId && !string.IsNullOrEmpty((string)Get(fresh, "Id")), "new profile has distinct identity");
        Check((int)Get(fresh, "TimeoutSeconds") == 120, "new profile timeout default");
        Check((string)Get(fresh, "Provider") == "openai", "new profile provider default");
        Check((string)Get(fresh, "PromptInput") == "question" && (string)Get(fresh, "RouteInput") == "route" && (string)Get(fresh, "AnswerOutput") == "answer", "Dify mapping defaults");
        Check(((System.Collections.IList)Get(Call(editor, "Snapshot"), "Profiles")).Count == 1, "new draft is not silently saved");
        Guid parsedId;
        Check(Guid.TryParseExact((string)Get(fresh, "Id"), "D", out parsedId), "new profile uses backend-required D-format identity");
        Set(fresh, "Name", "模型 B"); Set(fresh, "Endpoint", "https://example.com/v1/workflows/run");
        Set(fresh, "ApiKey", "synthetic-key-b"); Set(fresh, "Model", "test-model-b");
        Call(editor, "SaveProfile", fresh); string freshId = (string)Get(fresh, "Id");
        Check(((System.Collections.IList)Get(Call(editor, "Snapshot"), "Profiles")).Count == 2, "multiple profiles saved");
        Call(editor, "SetDefault", freshId);
        Check((string)Get(Call(editor, "Snapshot"), "DefaultProfileId") == freshId, "explicit default switch");
        Call(editor, "DeleteProfile", freshId);
        Check((string)Get(Call(editor, "Snapshot"), "DefaultProfileId") == originalId, "deleting default selects remaining profile");
        object snapshot = Call(editor, "Snapshot");
        Set(((System.Collections.IList)Get(snapshot, "Profiles"))[0], "Name", "snapshot mutation");
        Check((string)Get(Call(editor, "Find", originalId), "Name") == "修改 A", "persist snapshot is isolated");
        bool rejected = false; Set(copy, "Name", "");
        try { Call(editor, "SaveProfile", copy); } catch (TargetInvocationException) { rejected = true; }
        Check(rejected && (string)Get(Call(editor, "Find", originalId), "Name") == "修改 A", "invalid saves preserve prior profile");
        rejected = false; try { Call(editor, "SetDefault", "missing"); } catch (TargetInvocationException) { rejected = true; }
        Check(rejected, "unknown default rejected");
        Call(editor, "DeleteProfile", originalId);
        Check(string.IsNullOrEmpty((string)Get(Call(editor, "Snapshot"), "DefaultProfileId")), "empty list clears default");
        Set(profile, "Name", "搞快点 \u0043\u0065\u006c\u006c\u006d 兼容性测试");
        object migrated = Activator.CreateInstance(editorType, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public,
            null, new[] { settings }, null);
        Check((string)Get(Call(migrated, "Find", originalId), "Name") == "搞快点智擎工作流", "exact historic preset display name migrates");
        Check((string)Get(profile, "Name") != "搞快点智擎工作流", "name migration does not mutate loaded settings");
        Set(profile, "Name", "我的智擎配置");
        object custom = Activator.CreateInstance(editorType, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public,
            null, new[] { settings }, null);
        Check((string)Get(Call(custom, "Find", originalId), "Name") == "我的智擎配置", "custom names preserved");
    }
    static void ApiKeyVisibilityBehavior(SyntaxTree tree, System.Collections.Generic.IEnumerable<MetadataReference> refs)
    {
        var form = tree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().Single(c => c.Identifier.Text == "GkdAiSettingsForm");
        var method = form.Members.OfType<MethodDeclarationSyntax>().SingleOrDefault(m => m.Identifier.Text == "SetApiKeyVisible");
        Check(method != null, "key visibility action exists");
        string harness = @"class TextBox { public bool UseSystemPasswordChar = true; public string Text = ""synthetic-existing-key""; }
class Button { public string AccessibleName = ""显示 API Key""; public void Invalidate() {} }
class VisibilityHarness {
readonly TextBox apiKey = new TextBox(); readonly Button showApiKey = new Button(); bool dirty;
public bool Masked { get { return apiKey.UseSystemPasswordChar; } }
public string Key { get { return apiKey.Text; } }
public string Caption { get { return showApiKey.AccessibleName; } }
public bool Dirty { get { return dirty; } }
" + method.ToFullString() + "}";
        var compilation = CSharpCompilation.Create("ApiKeyVisibilityBehavior", new[] { CSharpSyntaxTree.ParseText(harness) }, refs,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        using (var stream = new MemoryStream())
        {
            var result = compilation.Emit(stream);
            Check(result.Success, "real key visibility method compiles: " + string.Join("; ", result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error)));
            object view = Activator.CreateInstance(Assembly.Load(stream.ToArray()).GetType("VisibilityHarness"));
            Check((bool)Get(view, "Masked") && (string)Get(view, "Caption") == "显示 API Key", "new view starts masked with visible reveal action");
            Call(view, "SetApiKeyVisible", true);
            Check(!(bool)Get(view, "Masked") && (string)Get(view, "Caption") == "隐藏 API Key", "reveal shows existing key with hide action");
            Check((string)Get(view, "Key") == "synthetic-existing-key" && !(bool)Get(view, "Dirty"), "revealing preserves existing key and does not dirty settings");
            Call(view, "SetApiKeyVisible", false);
            Check((bool)Get(view, "Masked") && (string)Get(view, "Caption") == "显示 API Key", "hide restores password masking and reveal action");
            Check((string)Get(view, "Key") == "synthetic-existing-key", "hide never erases or rewrites key");
        }
        string source = form.ToFullString();
        Check(source.Contains("PaintApiKeyEye") && source.Contains("g.DrawPath(pen, path)"), "drawn compact eye toggle");
        Check(form.DescendantNodes().OfType<LambdaExpressionSyntax>().Any(l => l.ToString().Contains("SetApiKeyVisible(apiKey.UseSystemPasswordChar)")), "click toggles current mask state");
        var constructor = form.Members.OfType<ConstructorDeclarationSyntax>().Single().ToFullString();
        Check(constructor.Contains("var keyRow = new Panel") && source.Contains("BorderStyle = BorderStyle.None") && source.Contains("Dock = DockStyle.Right"), "key and eye inside shared border");
        Check(constructor.Contains("keyRow.Controls.Add(apiKey)") && constructor.Contains("keyRow.Controls.Add(showApiKey)") && constructor.Contains("AddField(settingsPanel, \"API Key\", keyRow)"), "key and reveal share existing field width");
        var show = form.Members.OfType<MethodDeclarationSyntax>().Single(m => m.Identifier.Text == "ShowProfile").ToFullString();
        Check(show.Contains("SetApiKeyVisible(false)"), "profile switches reset key to hidden");
        Check(source.Contains("SetApiKeyVisible(false); Close()") && source.Contains("if (!e.Cancel) { SetApiKeyVisible(false);"), "close resets key to hidden");
    }
    static void CompileNativeForm(string source, string profileSource, string helpSource, string generalHelpSource)
    {
        string root = Environment.GetEnvironmentVariable("NUGET_PACKAGES");
        if (string.IsNullOrEmpty(root)) root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".nuget", "packages");
        string referencesPath = Path.Combine(root, "microsoft.netframework.referenceassemblies.net48", "1.0.3", "build", ".NETFramework", "v4.8");
        Check(Directory.Exists(referencesPath), "net48 reference assemblies available for native form compile");
        string formsSource = File.ReadAllText(Path.Combine(Path.GetDirectoryName(SourcePath), "Forms.cs"));
        var sharedUi = CSharpSyntaxTree.ParseText(formsSource).GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().Single(c => c.Identifier.Text == "Ui");
        string seams = @"using System; using System.Drawing; using System.Windows.Forms; using System.Threading; using System.Threading.Tasks;
namespace Outline.Addin {
internal static class GkdAiFunctions { internal static void ConfigurationChanged() {} }
internal static class GkdAiDiagnostics { internal static string LogDirectory { get { return ""synthetic-log-directory""; } } }
internal static class GkdAiStore { public static GkdAiSettings Load() { throw new NotSupportedException(); } public static void Save(GkdAiSettings settings) { throw new NotSupportedException(); } }
internal sealed class GkdAiReply { public object Value {get;set;} public long ElapsedMilliseconds {get;set;} }
internal static class GkdAiClient { public static Task<GkdAiReply> ExecuteAsync(GkdAiProfile profile, string prompt, object[,] data, string mode, CancellationToken token) { throw new NotSupportedException(); } }
" + sharedUi.ToFullString() + "}";
        var options = new CSharpParseOptions(LanguageVersion.CSharp7_3);
        var compilation = CSharpCompilation.Create("NativeFormCompileOnly", new[] { source, profileSource, helpSource, generalHelpSource, seams }.Select(s => CSharpSyntaxTree.ParseText(s, options)),
            new[] { "mscorlib.dll", "System.dll", "System.Core.dll", "System.Drawing.dll", "System.Windows.Forms.dll" }.Select(name => MetadataReference.CreateFromFile(Path.Combine(referencesPath, name))),
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        using (var stream = new MemoryStream())
        {
            var result = compilation.Emit(stream);
            Check(result.Success, "full native form compiles against net48 and existing Ui: " + string.Join("; ", result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error)));
        }
    }
    static void TestStatusBehavior(SyntaxTree tree, string profileSource, System.Collections.Generic.IEnumerable<MetadataReference> refs)
    {
        var form = tree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().Single(c => c.Identifier.Text == "GkdAiSettingsForm");
        var run = form.Members.OfType<MethodDeclarationSyntax>().Single(m => m.Identifier.Text == "RunTestAsync");
        var redact = form.Members.OfType<MethodDeclarationSyntax>().Single(m => m.Identifier.Text == "Redact");
        var probe = tree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().Single(c => c.Identifier.Text == "GkdAiSettingsProbe");
        string harness = @"using System.Linq; using System.Globalization; using System.IO; using System.Diagnostics; using System.Threading; using System.Threading.Tasks;
" + profileSource + @"
namespace Outline.Addin {
internal sealed class GkdAiReply { internal object Value; }
internal static class GkdAiClient {
 internal static System.Exception Error; internal static object Value = ""GKD_OK""; internal static int Calls;
 internal static Task<GkdAiReply> ExecuteAsync(GkdAiProfile profile, string prompt, object[,] data, string mode, CancellationToken token) {
  Calls++; return Error == null ? Task.FromResult(new GkdAiReply { Value = Value }) : Task.FromException<GkdAiReply>(Error);
 }
}
internal sealed class FakeStatus { internal string Text = """"; }
internal sealed class StatusHarness {
 internal GkdAiProfile Profile = new GkdAiProfile { Name = ""fixture"", ApiKey = ""synthetic-ui-key"" };
 internal readonly FakeStatus testStatus = new FakeStatus(); CancellationTokenSource cancellation;
 bool closing = false; bool IsDisposed = false; bool Disposing = false;
 internal bool Busy; internal string Status { get { return testStatus.Text; } }
 GkdAiProfile ReadDraft() { return Profile; } void SetBusy(bool busy) { Busy = busy; }
" + run.ToFullString() + redact.ToFullString() + "}" + probe.ToFullString() + "}";
        var compilation = CSharpCompilation.Create("RealTestStatusBehavior", new[] { CSharpSyntaxTree.ParseText(harness) }, refs,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        using (var stream = new MemoryStream())
        {
            var result = compilation.Emit(stream);
            Check(result.Success, "real async status action compiles: " + string.Join("; ", result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error)));
            var assembly = Assembly.Load(stream.ToArray());
            Type client = assembly.GetType("Outline.Addin.GkdAiClient");
            Type viewType = assembly.GetType("Outline.Addin.StatusHarness");
            object view = Activator.CreateInstance(viewType, true);
            Action<Exception, object, bool> execute = (error, value, table) =>
            {
                client.GetField("Error", BindingFlags.NonPublic | BindingFlags.Static).SetValue(null, error);
                client.GetField("Value", BindingFlags.NonPublic | BindingFlags.Static).SetValue(null, value);
                ((System.Threading.Tasks.Task)Call(view, "RunTestAsync", table)).GetAwaiter().GetResult();
                Check(!(bool)viewType.GetField("Busy", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(view), "busy state released after test");
            };
            Func<string> status = () => (string)viewType.GetProperty("Status", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(view);
            execute(new InvalidDataException("请求 r-fixture；阶段 response；类别 HTTP；状态 401"), null, false);
            Check(status().Contains("r-fixture") && status().Contains("401") && status().Contains("打开诊断日志"), "request diagnostics survive actual async catch path");
            execute(new TimeoutException("请求 r-timeout；阶段 send；类别 timeout"), null, false);
            Check(status().Contains("r-timeout") && status().Contains("测试超时"), "timeout safe diagnostics shown");
            execute(new InvalidOperationException("unsafe response synthetic-ui-key"), null, false);
            Check(!status().Contains("unsafe response") && !status().Contains("synthetic-ui-key") && status().Contains("未预期错误"), "arbitrary invalid-operation details not exposed");
            execute(new InvalidDataException("safe diagnostic synthetic-ui-key"), null, false);
            Check(!status().Contains("synthetic-ui-key") && status().Contains("[密钥已隐藏]"), "request detail defense-in-depth key masking");
            execute(null, "not a grid", true);
            Check(status().Contains("结果校验失败") && status().Contains("5 行 × 2 列"), "local probe failure gives actionable safe detail");
            execute(null, "GKD_OK", false);
            Check(status().Contains("测试通过") && status().Contains("GKD_OK"), "success shows actual scalar probe result");
            object profile = viewType.GetField("Profile", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(view);
            Set(profile, "ApiKey", "");
            int calls = (int)client.GetField("Calls", BindingFlags.NonPublic | BindingFlags.Static).GetValue(null);
            execute(null, "GKD_OK", false);
            Check(status().Contains("尚未发送请求") && status().Contains("API 密钥") && (int)client.GetField("Calls", BindingFlags.NonPublic | BindingFlags.Static).GetValue(null) == calls, "invalid draft shows exact validation detail without calling backend");
        }
    }
    static void UiContract(SyntaxTree tree, string helpSource, string generalHelpSource)
    {
        var form = tree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().SingleOrDefault(c => c.Identifier.Text == "GkdAiSettingsForm");
        Check(form != null && form.BaseList.ToString().Contains("Form"), "native Form entry point");
        Check(form.Members.OfType<ConstructorDeclarationSyntax>().Any(c => c.ParameterList.Parameters.Count == 0), "parameterless constructor");
        string source = form.ToFullString();
        foreach (string label in new[] { "新建", "保存配置", "删除配置", "设为默认", "连接 / 鉴权测试", "表格能力测试", "取消测试", "填入已验证智擎测试配置" })
            Check(source.Contains("Ui.Button(\"" + label + "\""), "clickable " + label);
        Check(form.DescendantNodes().OfType<LambdaExpressionSyntax>().Any(lambda => lambda.ToString().Contains("LoadVerifiedZhiqingTest")), "verified Dify preset action wired");
        Check(source.Contains("question") && source.Contains("route") && source.Contains("answer") && source.Contains("hdec235"), "preset maps verified Dify input/output fields and route");
        Check(source.Contains("搞快点智擎工作流") && source.Contains("粘贴该工作流应用自己的 API Key"), "preset uses product name and requires own key");
        string legacyBrand = "\u0043\u0065\u006c\u006c\u006d";
        Check(source.IndexOf(legacyBrand, StringComparison.OrdinalIgnoreCase) < 0 && helpSource.IndexOf(legacyBrand, StringComparison.OrdinalIgnoreCase) < 0, "settings and help use neutral product wording");
        Check(source.Contains("\"gkd-workflow.key\"") && source.Contains("if (!File.Exists(keyPath))") && source.Contains("\"gkd-test-workflow.key\""), "neutral credential file preferred with legacy fallback");
        Check(source.Contains("Ui.Base(this"), "shared native theme");
        Check(source.Contains("UseSystemPasswordChar = true"), "key masked");
        foreach (string field in new[] { "Name", "Provider", "Endpoint", "Model", "ApiKey", "TimeoutSeconds", "PromptInput", "RouteInput", "AnswerOutput" })
            Check(source.Contains("profile." + field), "editable profile binding " + field);
        var run = form.Members.OfType<MethodDeclarationSyntax>().Single(m => m.Identifier.Text == "RunTestAsync").ToFullString();
        Check(run.Contains("await GkdAiClient.ExecuteAsync(") && run.Contains("cancellation.Token"), "real async cancellable backend request");
        Check(!run.Contains("GkdAiStore.Save") && !run.Contains("SaveCurrent"), "tests never persist unsaved edits");
        Check(run.Contains("GkdAiSettingsProbe.ValidateReply(reply.Value, table)"), "tests inspect reply contents");
        Check(run.Contains("stopwatch.ElapsedMilliseconds") && run.Contains("yyyy-MM-dd HH:mm:ss"), "status has total latency and timestamp");
        Check(run.Contains("IsDisposed") && run.Contains("Disposing"), "continuation guards disposed controls");
        Check(source.Contains("OnFormClosing") && source.Contains("CancelTest();"), "closing cancels in-flight test");
        Check(source.Contains("settingsPanel.Enabled = !busy") && source.Contains("profiles.Enabled = !busy"), "conflicting edits disabled while running");
        Check(source.Contains("基于当前设置") && source.Contains("未保存"), "unsaved test scope visible");
        Check(helpSource.Contains("动态数组") && helpSource.Contains("空白") && helpSource.Contains("单轮"), "formula limitations in help dialog");
        Check(source.Contains("加载配置失败") && source.Contains("if (loadFailed)"), "load failure stays visible after draft display");
        var delete = form.Members.OfType<MethodDeclarationSyntax>().Single(m => m.Identifier.Text == "DeleteCurrent").ToFullString();
        Check(delete.Contains("Profiles.Count == 1") && delete.Contains("至少保留一个"), "last saved profile explains backend minimum");
        Check(source.Contains("Maximum = 600"), "timeout editor matches real backend limit");
        Check(helpSource.Contains("标题、空行、表头"), "table layout includes title in help dialog");
        Check(source.Contains("SizeType.Absolute, 170"), "profile action area fits all four native buttons");
        HelpAndDiagnosticsContract(form, helpSource, generalHelpSource);
    }
    static void HelpAndDiagnosticsContract(ClassDeclarationSyntax form, string helpSource, string generalHelpSource)
    {
        string source = form.ToFullString();
        var constructor = form.Members.OfType<ConstructorDeclarationSyntax>().Single().ToFullString();
        Check(source.Contains("Ui.Button(\"★ 使用帮助与示例\")") && constructor.Contains("new GkdAiHelpForm()") && constructor.Contains("dialog.ShowDialog(this)"), "star help opens owned native dialog");
        Check(!constructor.Contains("=GKD(") && !constructor.Contains("不自动追加路径"), "settings page no longer embeds long usage guide");
        var helpTree = CSharpSyntaxTree.ParseText(helpSource, new CSharpParseOptions(LanguageVersion.CSharp7_3));
        Check(!helpTree.GetDiagnostics().Any(d => d.Severity == DiagnosticSeverity.Error), "help is C# 7.3 syntax");
        Check(helpSource.Contains("ReadOnly = true") && helpSource.Contains("document.SelectAll()") && helpSource.Contains("按 Ctrl+C 复制"), "examples selectable and copyable");
        Check(helpSource.Contains("RichTextBox") && helpSource.Contains("WordWrap = true") && helpSource.Contains("DetectUrls = false"), "guide renders as one wrapped plain-text document");
        Check(helpSource.Contains("ScrollBars = RichTextBoxScrollBars.Vertical") && !helpSource.Contains("AutoScroll = true") && !helpSource.Contains("content.SizeChanged") && !helpSource.Contains("new TextBox"), "one scroll host without child scrollbars or per-block layout");
        Check(helpSource.Contains("BeginInvoke") && helpSource.Contains("SelectionStart = 0") && helpSource.Contains("ScrollToCaret"), "AI guide resets scroll position after first layout");
        Check(generalHelpSource.Contains("Shown +=") && generalHelpSource.Contains("SelectionStart = 0") && generalHelpSource.Contains("ScrollToCaret"), "general help opens at top");
        Check(!generalHelpSource.Contains("content.SizeChanged"), "general help has no per-section resize handlers");
        Check(!helpSource.Contains("WebView") && !helpSource.Contains("WebBrowser"), "help has no embedded browser");
        var examples = helpTree.GetRoot().DescendantNodes().OfType<InvocationExpressionSyntax>()
            .Where(i => i.Expression.ToString() == "AddCode")
            .Select(i => ((LiteralExpressionSyntax)i.ArgumentList.Arguments[0].Expression).Token.ValueText).ToArray();
        Check(examples.Any(s => s.StartsWith("=A2&CHAR(10)&GKD(") && s.Contains("俄语") && s.Contains(",A2)")), "Chinese first newline Russian translation formula");
        Check(helpSource.Contains("填充柄") && helpSource.Contains("A3、A4"), "drag fill explains relative references");
        Check(examples.Any(s => s == "=GKD(\"翻译成俄语，只返回译文\",A2)"), "plain translation formula");
        Check(examples.Any(s => s.Contains("π*(D-t)*t*L*ρ/1000000") && s.Contains("单位kg")) && helpSource.Contains("适用标准"), "pipe mass formula specifies units and standards caveats");
        Check(examples.Any(s => s.Contains("定额参考") && s.Contains("禁止编造编码")), "quota recommendation bounded by supplied reference");
        Check(examples.Any(s => s.Contains("给定报价数据") && s.Contains("不补造实时价格")), "market price uses supplied quotes");
        Check(examples.Any(s => s.Contains("给定银行牌价")) && examples.Any(s => s.Contains("后端联网浏览") && s.Contains("禁止猜测")), "separate supplied FX and backend URL browsing examples");
        Check(helpSource.Contains("GKD 加载项自身不会抓取网址"), "URL does not promise addin browsing");
        Check(examples.Any(s => s.Contains("给定项目来源") && s.EndsWith(",A1:F30,1)")), "investment table grounded in project sources");
        Check(examples.Any(s => s.Contains(",A2,0)")) && examples.Any(s => s.EndsWith(",A2,1)")) && helpSource.Contains("第三参数必须填 1"), "table flag distinguished from omitted or zero scalar");
        Check(examples.All(s => s.Split(new[] { "\r\n" }, StringSplitOptions.None).All(line => line.StartsWith("=GKD(") || line.StartsWith("=PIPE_MASS(") || line.StartsWith("=GKD_CUSTOM(") || line.StartsWith("=A2&CHAR(10)&GKD(") || line.StartsWith("=PI()"))) && !helpSource.Contains("=搞快点("), "formulas use actual GKD or declared custom formula registration");
        var logs = form.Members.OfType<MethodDeclarationSyntax>().Single(m => m.Identifier.Text == "OpenDiagnosticsDirectory").ToFullString();
        Check(constructor.Contains("openLogs.Click +=") && logs.Contains("GkdAiDiagnostics.LogDirectory") && logs.Contains("Directory.CreateDirectory(directory)") && logs.Contains("UseShellExecute = true") && logs.Contains("catch (Exception)"), "log button creates and shell opens guarded diagnostics directory");
        var run = form.Members.OfType<MethodDeclarationSyntax>().Single(m => m.Identifier.Text == "RunTestAsync");
        foreach (string type in new[] { "InvalidDataException", "TimeoutException", "ArgumentException" })
        {
            var clause = run.DescendantNodes().OfType<CatchClauseSyntax>().Single(c => c.Declaration != null && c.Declaration.Type.ToString() == type);
            Check(clause.ToString().Contains("Redact(ex.Message, profile)"), "safe detailed " + type + " visible");
        }
        var generic = run.DescendantNodes().OfType<CatchClauseSyntax>().Single(c => c.Declaration != null && c.Declaration.Type.ToString() == "Exception");
        Check(!generic.ToString().Contains(".Message"), "unexpected exception text stays sanitized");
        var probeCall = run.DescendantNodes().OfType<TryStatementSyntax>().Single(t => t.Block.ToString().Contains("GkdAiSettingsProbe.ValidateReply") && !t.Block.ToString().Contains("ExecuteAsync"));
        Check(probeCall.Catches.Single().Declaration.Type.ToString() == "InvalidOperationException", "only local probe details exposed for invalid operation");
    }
    static void Reject(Action action, string message)
    {
        bool rejected = false; try { action(); } catch (TargetInvocationException) { rejected = true; }
        Check(rejected, message);
    }
    static object Call(object value, string method, params object[] args) { return value.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic).Invoke(value, args); }
    static object Get(object value, string name) { return value.GetType().GetProperty(name).GetValue(value); }
    static void Set(object value, string name, object data) { value.GetType().GetProperty(name).SetValue(value, data); }
    static void Check(bool condition, string message) { if (!condition) throw new Exception("FAIL " + message); checks++; }
}
