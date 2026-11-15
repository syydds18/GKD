using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

static partial class Program
{
    static int checks;
    static readonly string SourceDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../src/Outline.Addin"));
    static void Main()
    {
        string source = File.ReadAllText(Path.Combine(SourceDir, "GkdCustomFormulaForm.cs"));
        var tree = CSharpSyntaxTree.ParseText(source, new CSharpParseOptions(LanguageVersion.CSharp7_3));
        Check(!tree.GetDiagnostics().Any(d => d.Severity == DiagnosticSeverity.Error), "C# 7.3 syntax");
        CompileNative(source);
        Behavior(tree);
        AsyncBehavior(tree);
        Console.WriteLine("PASS " + checks + " custom formula UI compile, behavior and safety checks; Windows GUI not exercised.");
    }
    static void CompileNative(string source)
    {
        string packageRoot = Environment.GetEnvironmentVariable("NUGET_PACKAGES");
        if (string.IsNullOrEmpty(packageRoot)) packageRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".nuget", "packages");
        string referenceRoot = Path.Combine(packageRoot, "microsoft.netframework.referenceassemblies.net48", "1.0.3", "build", ".NETFramework", "v4.8");
        Check(Directory.Exists(referenceRoot), "net48 references available");
        var ui = CSharpSyntaxTree.ParseText(File.ReadAllText(Path.Combine(SourceDir, "Forms.cs"))).GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().Single(x => x.Identifier.Text == "Ui");
        string seams = @"using System; using System.Collections.Generic; using System.Threading; using System.Threading.Tasks; using System.Drawing; using System.Windows.Forms;
namespace Outline.Addin {
internal sealed class GkdAiReply { public object Value {get;set;} }
internal static class GkdAiStore { internal static GkdAiSettings Load() { throw new NotSupportedException(); } internal static GkdAiProfile ResolveDefault(GkdAiSettings settings) { throw new NotSupportedException(); } }
internal static class GkdAiClient { internal static Task<GkdAiReply> ExecuteAsync(GkdAiProfile p, string prompt, object[,] data, string mode, CancellationToken token) { throw new NotSupportedException(); } }
internal static class GkdCustomFormulaStore {
 internal static List<GkdCustomFormula> Load() { throw new NotSupportedException(); }
 internal static void Save(List<GkdCustomFormula> items) { throw new NotSupportedException(); }
 internal static List<GkdCustomFormula> Parse(string json) { throw new NotSupportedException(); }
 internal static string Serialize(List<GkdCustomFormula> items) { throw new NotSupportedException(); }
}
internal static class GkdCustomFormulaFunctions { internal static void Reload() { throw new NotSupportedException(); } }
internal static class GkdCustomFormulaRuntime { internal static string DescriptionText = string.Empty; }
" + ui.ToFullString() + "}";
        string jsonDll = Path.Combine(packageRoot, "newtonsoft.json", "13.0.3", "lib", "net45", "Newtonsoft.Json.dll");
        var refs = new[] { "mscorlib.dll", "System.dll", "System.Core.dll", "System.Drawing.dll", "System.Windows.Forms.dll" }.Select(x => MetadataReference.CreateFromFile(Path.Combine(referenceRoot, x))).Concat(new[] { MetadataReference.CreateFromFile(jsonDll) });
        Compile("NativeCustomFormulaUi", new[] { source, File.ReadAllText(Path.Combine(SourceDir, "GkdCustomFormula.cs")), File.ReadAllText(Path.Combine(SourceDir, "GkdAiProtocol.cs")), File.ReadAllText(Path.Combine(SourceDir, "GkdAiProfile.cs")), seams }, refs, false);
        Check(true, "complete form compiles against actual model, shared Ui and net48 WinForms");
    }
    static void Behavior(SyntaxTree tree)
    {
        var root = tree.GetRoot();
        var form = root.DescendantNodes().OfType<ClassDeclarationSyntax>().Single(x => x.Identifier.Text == "GkdCustomFormulaForm");
        var editor = root.DescendantNodes().OfType<ClassDeclarationSyntax>().Single(x => x.Identifier.Text == "GkdCustomFormulaEditorForm");
        var methods = form.Members.OfType<MethodDeclarationSyntax>().Where(x => new[] { "Copy", "Describe", "SaveDraft", "ValidateForReview" }.Contains(x.Identifier.Text)).Select(x => x.ToFullString());
        var save = editor.Members.OfType<MethodDeclarationSyntax>().Single(x => x.Identifier.Text == "SaveCurrent");
        var prompt = root.DescendantNodes().OfType<ClassDeclarationSyntax>().Single(x => x.Identifier.Text == "GkdCustomFormulaDesignPrompt");
        string trial = root.DescendantNodes().OfType<ClassDeclarationSyntax>().SingleOrDefault(x => x.Identifier.Text == "GkdCustomFormulaTrial")?.ToFullString() ?? "";
        string harness = @"using System; using System.Collections.Generic; using System.Linq; using System.IO; using System.Globalization;
namespace Outline.Addin {
internal enum DialogResult { None, OK, Cancel }
internal static class GkdCustomFormulaStore { internal static int Writes; internal static bool Fail; internal static List<GkdCustomFormula> Items;
 internal static void Save(List<GkdCustomFormula> items) { if(Fail) throw new IOException(); Writes++; Items = items; } }
internal sealed class GkdCustomFormulaPreviewForm : IDisposable { internal static DialogResult Next; internal static int Seen; internal GkdCustomFormulaPreviewForm(string a,string b,bool c){Seen++;} internal DialogResult ShowDialog(object owner){return Next;} public void Dispose(){} }
internal sealed class FakeStatus { internal string Text; }
internal static class GkdCustomFormulaFunctions { internal static void Reload(){} }
internal sealed class GkdCustomFormulaForm { internal List<GkdCustomFormula> saved = new List<GkdCustomFormula>(); internal void RefreshList(string id){} " + string.Join("\n", methods) + @" }
internal sealed class EditorHarness { internal object cancellation; internal GkdCustomFormula Draft; internal Action<GkdCustomFormula> persist; internal bool dirty = true; internal DialogResult DialogResult; internal FakeStatus status = new FakeStatus(); internal bool Closed;
 internal GkdCustomFormula ReadDraft(){return Draft;} internal void Close(){Closed=true;} " + save.ToFullString() + " }" + prompt.ToFullString() + trial + "}";
        var refs = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")).Split(Path.PathSeparator).Select(x => MetadataReference.CreateFromFile(x)).Concat(new[] { MetadataReference.CreateFromFile(typeof(Newtonsoft.Json.JsonConvert).Assembly.Location) });
        var assembly = Compile("CustomFormulaUiBehavior", new[] { harness, File.ReadAllText(Path.Combine(SourceDir, "GkdCustomFormula.cs")), File.ReadAllText(Path.Combine(SourceDir, "GkdAiProtocol.cs")) }, refs, true);
        Type model = assembly.GetType("Outline.Addin.GkdCustomFormula");
        object draft = Activator.CreateInstance(model, true);
        Set(draft, "ExcelName", "PIPE"); Set(draft, "DisplayName", "钢管质量"); Set(draft, "Description", "D 外径，单位 mm；T 壁厚，单位 mm；输出 kg/m"); Set(draft, "FormulaExpression", "3.14159*(D-T)*T*7850/1000000");
        ((System.Collections.IList)Get(draft, "Inputs")).Add("D"); ((System.Collections.IList)Get(draft, "Inputs")).Add("T");
        Type formType = assembly.GetType("Outline.Addin.GkdCustomFormulaForm");
        object copy = CallStatic(formType, "Copy", draft);
        ((System.Collections.IList)Get(copy, "Inputs")).Add("OTHER");
        Check(((System.Collections.IList)Get(draft, "Inputs")).Count == 2, "UI draft copy deeply isolates input names");
        Set(copy, "Description", "");
        Reject(() => CallStatic(formType, "ValidateForReview", copy), "UI rejects definitions without meaning explanation");
        string description = (string)CallStatic(formType, "Describe", draft);
        foreach (string value in new[] { "钢管质量", "PIPE", "单位 mm", "3.14159", "D, T" }) Check(description.Contains(value), "meaning preview includes " + value);
        Check(description.Contains("=PIPE(A2,B2)") && description.Contains("=GKD_CUSTOM(\"PIPE\",A2,B2)"), "preview provides direct and generic Excel call examples");
        object constant = CallStatic(formType, "Copy", draft);
        ((System.Collections.IList)Get(constant, "Inputs")).Clear(); Set(constant, "FormulaExpression", "1");
        string zeroDescription = (string)CallStatic(formType, "Describe", constant);
        Check(zeroDescription.Contains("=PIPE()") && zeroDescription.Contains("=GKD_CUSTOM(\"PIPE\")"), "zero input examples omit argument separator");
        string managerSource = form.ToFullString();
        Check(managerSource.Contains("formulas.Format +=") && managerSource.Contains("FormattingEnabled = true") && managerSource.Contains("DisplayMember = \"DisplayName\""), "list formatting preserves bound formula objects and display member");
        Type trialType = assembly.GetType("Outline.Addin.GkdCustomFormulaTrial");
        Check(trialType != null, "numeric sample trial helper exists");
        object division = CallStatic(formType, "Copy", draft);
        Set(division, "FormulaExpression", "D/T"); Set(division, "OutputUnit", "m");
        string sampleResult = (string)CallStatic(trialType, "Evaluate", division, "12, 3");
        Check(sampleResult.Contains("4") && sampleResult.Contains("m") && sampleResult.Contains("不能证明算法正确"), "trial uses declared argument order and includes result unit and sample limitation");
        foreach (string invalid in new[] { "", "12", "12,3,4", "NaN,3", "Infinity,3", "1e309,3", "12,,3", "12,nope", "12,0" })
            Reject(() => CallStatic(trialType, "Evaluate", division, invalid), "trial rejects invalid count/value/arithmetic: " + invalid);
        Set(division, "FormulaExpression", "");
        Reject(() => CallStatic(trialType, "Evaluate", division, "12,3"), "blank new algorithm never defaults to zero");
        Check(((string)CallStatic(trialType, "Evaluate", constant, "")).Contains("1"), "zero-input trial evaluates actual constant");
        Set(constant, "FormulaExpression", "CONVERT(1000,\"mm\",\"m\")");
        Check(((string)CallStatic(trialType, "Evaluate", constant, "")).Contains("1"), "trial uses actual restricted CONVERT evaluator");
        var promptType = assembly.GetType("Outline.Addin.GkdCustomFormulaDesignPrompt");
        string designPrompt = (string)CallStatic(promptType, "Build", draft, "计算延米质量");
        Check(designPrompt.Contains("计算延米质量") && designPrompt.Contains("不要 Markdown") && designPrompt.Contains("FormulaExpression"), "design prompt includes request, strict schema and data-only boundary");
        Check(designPrompt.Contains("CONVERT") && designPrompt.Contains("3.141592653589793") && !designPrompt.Contains("禁止函数调用。"), "design prompt documents restricted CONVERT and numeric PI constant");
        string editorSource = editor.ToFullString();
        Check(editorSource.Contains("CONVERT") && editorSource.Contains("3.141592653589793"), "editor documents supported grammar with numeric PI constant");
        Check(editorSource.Contains("试算") && editorSource.Contains("样例值") && editorSource.Contains("TrialAsync"), "editor exposes numeric sample input and trial action");
        Check(editor.Members.OfType<FieldDeclarationSyntax>().Any(x => x.Declaration.Type.ToString() == "TextBox" && x.Declaration.Variables.Any(v => v.Identifier.Text == "status")), "long AI results and request diagnostics remain scrollable and selectable");
        int schema = designPrompt.IndexOf("{\"SchemaVersion\"");
        string json = designPrompt.Substring(schema).Split(new[] { "\n用户需求" }, StringSplitOptions.None)[0];
        var imported = (System.Collections.IList)CallStatic(assembly.GetType("Outline.Addin.GkdCustomFormulaJson"), "Import", json);
        Check(imported.Count == 1 && (string)Get(imported[0], "ExcelName") == "PIPE", "AI prompt schema round-trips through real production JSON importer");
        string wrappedPrompt = (string)CallStatic(assembly.GetType("Outline.Addin.GkdAiProtocol"), "BuildPrompt", designPrompt, null, "单值");
        Check(wrappedPrompt.Contains("data") && wrappedPrompt.Contains("null"), "design prompt accepted by real no-data protocol path");
        Reject(() => CallStatic(promptType, "Build", draft, ""), "blank AI request rejected");
        Reject(() => CallStatic(promptType, "Build", draft, new string('x', 8001)), "oversized AI request rejected");
        object manager = Activator.CreateInstance(formType, true);
        var action = Delegate.CreateDelegate(typeof(Action<>).MakeGenericType(model), manager, formType.GetMethod("SaveDraft", Flags));
        Type editorType = assembly.GetType("Outline.Addin.EditorHarness");
        object view = Activator.CreateInstance(editorType, true);
        SetField(view, "Draft", draft); SetField(view, "persist", action);
        Type previewType = assembly.GetType("Outline.Addin.GkdCustomFormulaPreviewForm");
        Type storeType = assembly.GetType("Outline.Addin.GkdCustomFormulaStore");
        Type dialogType = assembly.GetType("Outline.Addin.DialogResult");
        previewType.GetField("Next", Flags).SetValue(null, Enum.Parse(dialogType, "Cancel"));
        Call(view, "SaveCurrent");
        Check((int)storeType.GetField("Writes", Flags).GetValue(null) == 0 && !(bool)GetField(view, "Closed"), "cancelled preview does not persist or close draft");
        previewType.GetField("Next", Flags).SetValue(null, Enum.Parse(dialogType, "OK"));
        string algorithm = (string)Get(draft, "FormulaExpression"); Set(draft, "FormulaExpression", "");
        Call(view, "SaveCurrent");
        Check((int)storeType.GetField("Writes", Flags).GetValue(null) == 0 && !(bool)GetField(view, "Closed"), "blank expression cannot be confirmed into saved algorithm");
        Set(draft, "FormulaExpression", algorithm);
        storeType.GetField("Fail", Flags).SetValue(null, true);
        Call(view, "SaveCurrent");
        Check((int)storeType.GetField("Writes", Flags).GetValue(null) == 0 && !(bool)GetField(view, "Closed"), "failed save preserves open draft and original list");
        Check(((System.Collections.IList)GetField(manager, "saved")).Count == 0, "failed write does not publish pending manager state");
        storeType.GetField("Fail", Flags).SetValue(null, false);
        Call(view, "SaveCurrent");
        Check((int)storeType.GetField("Writes", Flags).GetValue(null) == 1 && (bool)GetField(view, "Closed"), "confirmed preview persists once and closes editor");
        Set(draft, "DisplayName", "later mutation");
        var stored = (System.Collections.IList)storeType.GetField("Items", Flags).GetValue(null);
        Check((string)Get(stored[0], "DisplayName") == "钢管质量", "stored candidate isolated from subsequent draft mutation");
        string previewSource = root.DescendantNodes().OfType<ClassDeclarationSyntax>().Single(x => x.Identifier.Text == "GkdCustomFormulaPreviewForm").ToFullString();
        Check(previewSource.Contains("Checked = false") && previewSource.Contains("confirm.Enabled = false") && previewSource.Contains("if (!reviewed.Checked) return"), "confirmation starts unchecked and guards click handler");
        string ai = editor.Members.OfType<MethodDeclarationSyntax>().Single(x => x.Identifier.Text == "DesignAsync").ToFullString();
        Check(ai.IndexOf("preview.ShowDialog") < ai.IndexOf("ExecuteAsync"), "AI request requires prompt preview approval");
        Check(!ai.Contains("persist(") && !ai.Contains("GkdCustomFormulaStore.Save"), "AI completion cannot save definitions");
        Check(ai.Contains("ThrowIfCancellationRequested") && ai.Contains("closing || IsDisposed || Disposing"), "cancelled and closed AI callbacks cannot publish drafts");
    }
    static Assembly Compile(string name, string[] sources, System.Collections.Generic.IEnumerable<MetadataReference> refs, bool load)
    {
        var compilation = CSharpCompilation.Create(name, sources.Select(x => CSharpSyntaxTree.ParseText(x, new CSharpParseOptions(LanguageVersion.CSharp7_3))), refs, new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        using (var stream = new MemoryStream()) { var result = compilation.Emit(stream); Check(result.Success, name + " compilation: " + string.Join("; ", result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error))); return load ? Assembly.Load(stream.ToArray()) : null; }
    }
    const BindingFlags Flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;
    static object Call(object o, string n, params object[] args) { return o.GetType().GetMethod(n, Flags).Invoke(o, args); }
    static object CallStatic(Type t, string n, params object[] args) { return t.GetMethod(n, Flags).Invoke(null, args); }
    static object Get(object o, string n) { return o.GetType().GetProperty(n).GetValue(o); }
    static void Set(object o, string n, object v) { o.GetType().GetProperty(n).SetValue(o, v); }
    static object GetField(object o, string n) { return o.GetType().GetField(n, Flags).GetValue(o); }
    static void SetField(object o, string n, object v) { o.GetType().GetField(n, Flags).SetValue(o, v); }
    static void Reject(Action action, string message) { bool failed = false; try { action(); } catch (TargetInvocationException) { failed = true; } Check(failed, message); }
    static void Check(bool condition, string message) { if (!condition) throw new Exception("FAIL " + message); checks++; }
}
