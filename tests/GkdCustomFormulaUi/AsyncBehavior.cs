using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;

static partial class Program
{
    static void AsyncBehavior(SyntaxTree tree)
    {
        var root = tree.GetRoot();
        var editor = root.DescendantNodes().OfType<ClassDeclarationSyntax>().Single(x => x.Identifier.Text == "GkdCustomFormulaEditorForm");
        Check(editor.Members.OfType<MethodDeclarationSyntax>().Any(x => x.Identifier.Text == "TrialAsync"), "asynchronous trial action exists");
        string methods = string.Join("\n", editor.Members.OfType<MethodDeclarationSyntax>().Where(x => new[] { "TrialAsync", "DesignAsync" }.Contains(x.Identifier.Text)).Select(x => x.ToFullString()));
        string helpers = string.Join("\n", root.DescendantNodes().OfType<ClassDeclarationSyntax>().Where(x => new[] { "GkdCustomFormulaTrial", "GkdCustomFormulaDesignPrompt" }.Contains(x.Identifier.Text)).Select(x => x.ToFullString()));
        string harness = @"using System; using System.Linq; using System.IO; using System.Collections.Generic; using System.Globalization; using System.Threading; using System.Threading.Tasks;
namespace Outline.Addin {
internal enum DialogResult { OK, Cancel }
internal sealed class FakeStatus { internal string Text = string.Empty; }
internal sealed class GkdAiReply { public object Value {get;set;} }
internal static class GkdAiStore { internal static GkdAiSettings Settings; internal static GkdAiSettings Load(){return Settings;} internal static GkdAiProfile ResolveDefault(GkdAiSettings settings){return settings.Profiles[0].Clone();} }
internal static class GkdAiClient { internal static int Calls; internal static GkdAiProfile Profile; internal static string Prompt; internal static object[,] Data; internal static TaskCompletionSource<GkdAiReply> Pending;
 internal static Task<GkdAiReply> ExecuteAsync(GkdAiProfile p,string prompt,object[,] data,string mode,CancellationToken token){Calls++;Profile=p;Prompt=prompt;Data=data;Pending=new TaskCompletionSource<GkdAiReply>();return Pending.Task;} }
internal sealed class GkdCustomFormulaPreviewForm : IDisposable { internal static DialogResult Next; internal static string Content;
 internal GkdCustomFormulaPreviewForm(string title,string content,bool saving){Content=content;} internal DialogResult ShowDialog(object o){return Next;} public void Dispose(){} }
internal static class GkdCustomFormulaForm { internal static void ValidateForReview(GkdCustomFormula f){f.Validate();} }
internal sealed class AsyncEditor { internal CancellationTokenSource cancellation; internal FakeStatus samples = new FakeStatus(), status = new FakeStatus(), request = new FakeStatus(); internal GkdCustomFormula Draft; internal bool closing, IsDisposed, Disposing, dirty, Busy;
 internal GkdCustomFormula ReadDraft(){return Draft;} internal void SetDraftFields(GkdCustomFormula draft){Draft=draft;} internal void SetBusy(bool value){Busy=value;}
" + methods + "}" + helpers + "}";
        var refs = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")).Split(Path.PathSeparator).Select(x => MetadataReference.CreateFromFile(x)).Concat(new[] { MetadataReference.CreateFromFile(typeof(Newtonsoft.Json.JsonConvert).Assembly.Location) });
        var assembly = Compile("CustomFormulaAsyncUi", new[] { harness, File.ReadAllText(Path.Combine(SourceDir, "GkdCustomFormula.cs")), File.ReadAllText(Path.Combine(SourceDir, "GkdAiProtocol.cs")), File.ReadAllText(Path.Combine(SourceDir, "GkdAiProfile.cs")) }, refs, true);
        var client = assembly.GetType("Outline.Addin.GkdAiClient"); var preview = assembly.GetType("Outline.Addin.GkdCustomFormulaPreviewForm");
        object profile = Activator.CreateInstance(assembly.GetType("Outline.Addin.GkdAiProfile"), true); Set(profile,"ApiKey","test-only-secret");
        object settings = Activator.CreateInstance(assembly.GetType("Outline.Addin.GkdAiSettings"), true); Set(settings,"DefaultProfileId",Get(profile,"Id")); ((System.Collections.IList)Get(settings,"Profiles")).Add(profile);
        assembly.GetType("Outline.Addin.GkdAiStore").GetField("Settings",Flags).SetValue(null,settings);
        object view = Activator.CreateInstance(assembly.GetType("Outline.Addin.AsyncEditor"), true);
        object draft = Activator.CreateInstance(assembly.GetType("Outline.Addin.GkdCustomFormula"), true);
        Set(draft,"ExcelName","AI_SAMPLE");Set(draft,"DisplayName","样例");Set(draft,"Description","参数单位 m");Set(draft,"OutputUnit","m");Set(draft,"PromptTemplate","根据输入计算");Set(draft,"Kind",Enum.Parse(assembly.GetType("Outline.Addin.GkdCustomFormulaKind"),"Ai"));((System.Collections.IList)Get(draft,"Inputs")).Add("D");
        SetField(view,"Draft",draft); SetField(GetField(view,"samples"),"Text","12");
        var dialog = assembly.GetType("Outline.Addin.DialogResult"); preview.GetField("Next",Flags).SetValue(null,Enum.Parse(dialog,"Cancel"));
        ((Task)Call(view,"TrialAsync")).GetAwaiter().GetResult(); Check((int)client.GetField("Calls",Flags).GetValue(null)==0,"cancelled trial preview sends no request");
        preview.GetField("Next",Flags).SetValue(null,Enum.Parse(dialog,"OK"));
        var task = (Task)Call(view,"TrialAsync");
        Check((bool)GetField(view,"Busy") && !(bool)GetField(view,"dirty"),"AI trial disables editing without changing draft");
        Check((string)client.GetField("Prompt",Flags).GetValue(null)=="根据输入计算" && !ReferenceEquals(client.GetField("Profile",Flags).GetValue(null),profile),"AI trial sends current prompt and cloned default profile");
        object[,] data = (object[,])client.GetField("Data",Flags).GetValue(null);
        Check(data.GetLength(0)==1 && (string)data[0,0]=="D" && (double)data[0,1]==12,"AI trial sends labelled ordered numeric arguments");
        Check(((string)preview.GetField("Content",Flags).GetValue(null)).Contains("D") && !((string)preview.GetField("Content",Flags).GetValue(null)).Contains("test-only-secret"),"preview shows exact labelled data without key");
        CompleteReply(client,assembly,"answer");task.GetAwaiter().GetResult();
        Check(Status(view).Contains("answer") && Status(view).Contains("m") && Status(view).Contains("不能证明算法正确") && !(bool)GetField(view,"Busy"),"AI trial shows result unit and limitation then resets busy state");
        task=(Task)Call(view,"TrialAsync"); ((CancellationTokenSource)GetField(view,"cancellation")).Cancel();CompleteReply(client,assembly,"late");task.GetAwaiter().GetResult();
        Check(Status(view).Contains("取消") && !Status(view).Contains("late"),"cancelled AI trial suppresses late reply");
        task=(Task)Call(view,"TrialAsync");SetField(view,"closing",true);SetField(GetField(view,"status"),"Text","closed sentinel");CompleteReply(client,assembly,"late");task.GetAwaiter().GetResult();Check(Status(view)=="closed sentinel","closed AI trial cannot publish reply");SetField(view,"closing",false);
        foreach(string method in new[]{"TrialAsync","DesignAsync"})
        {
            SetField(GetField(view,"request"),"Text","生成定义");
            foreach(Exception error in new Exception[]{new InvalidDataException("AI 请求失败：解析响应；编号 abcdef123456；详情见诊断日志。"),new TimeoutException("AI 请求失败：请求超时；编号 fedcba654321。")})
            {
                task=(Task)Call(view,method);object pending=client.GetField("Pending",Flags).GetValue(null);pending.GetType().GetMethod("SetException",new[]{typeof(Exception)}).Invoke(pending,new object[]{error});task.GetAwaiter().GetResult();Check(Status(view).Contains(error.Message) && !Status(view).Contains("test-only-secret"),method+" retains safe request diagnostic detail");
            }
            task=(Task)Call(view,method);object finalPending=client.GetField("Pending",Flags).GetValue(null);finalPending.GetType().GetMethod("SetException",new[]{typeof(Exception)}).Invoke(finalPending,new object[]{new Exception("test-only-secret")});task.GetAwaiter().GetResult();Check(!Status(view).Contains("test-only-secret"),method+" does not surface arbitrary exception text");
        }
    }
    static string Status(object view){return (string)GetField(GetField(view,"status"),"Text");}
    static void CompleteReply(Type client, Assembly assembly, string text){object reply=Activator.CreateInstance(assembly.GetType("Outline.Addin.GkdAiReply"),true);Set(reply,"Value",text);Call(client.GetField("Pending",Flags).GetValue(null),"SetResult",reply);}
}
