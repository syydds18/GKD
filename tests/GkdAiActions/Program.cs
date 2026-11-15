using System;
using System.Collections.Generic;
using System.Reflection;
using Outline.Addin;
using ExcelDna.Integration;

class Program
{
    static Type Actions { get { return typeof(Program).Assembly.GetType("Outline.Addin.GkdAiActions"); } }
    public static void Check(bool ok, string message) { if (!ok) throw new Exception(message); }
    static object Invoke(string name, params object[] args)
    {
        Check(Actions != null, "GKD native selection actions missing");
        var method = Actions.GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public);
        Check(method != null, "GKD action method missing: " + name);
        try { return method.Invoke(null, args); }
        catch (TargetInvocationException e) { throw e.InnerException; }
    }
    static void Main()
    {
        GkdCustomFormulaFunctions.SetKind("AI_WORKFLOW",true);
        GkdCustomFormulaFunctions.SetKind("OLD_AI_ALIAS",true);
        GkdCustomFormulaFunctions.SetKind("DETERMINISTIC_ALIAS",false);
        string[] yes = { "=GKD(A1,B1)", " =gkd (A1,B1)", "=_xll.GKD(A1,B1)", "=@GKD(A1,B1)", "=GKD_CUSTOM(\"AI_WORKFLOW\",A1)", "=AI_WORKFLOW(A1)", "=_xll.OLD_AI_ALIAS(A1)" };
        string[] no = { "=GKD.PROMPT(A1)", "=SUM(GKD(A1,B1))", "=IF(TRUE,GKD(A1,B1),0)", "=GKD2(A1)", "=Sheet1!GKD(A1)", "=\"GKD(A1)\"", "GKD(A1)", "=GKD(A1)+1", "=GKD(A1)&GKD(B1)", "=GKD(A1", "=DETERMINISTIC_ALIAS(A1)", "=GKD_CUSTOM(\"DETERMINISTIC_ALIAS\",A1)", "=NOT_REGISTERED(A1)" };
        foreach (var s in yes) Check((bool)Invoke("IsGkdFormula", s), "accept exact GKD call: " + s);
        foreach (var s in no) Check(!(bool)Invoke("IsGkdFormula", s), "reject unrelated formula: " + s);
        Check((bool)Invoke("IsGkdFormula", "=GKD(\"literal ) and \"\"quotes\"\"\",A1)"), "quoted parentheses do not end call");
        Console.WriteLine("PASS exact GKD formula recognition");
        RefreshSelection();
        FreezeSelection();
        FreezeSingleScalarWithoutSpill();
        FreezeGuards();
        FreezeFailureRestoresEvents();
        ConfirmationCannotRedirectWrite();
        MixedSelectionSkipsNonGkdCells();
    }
    static void RefreshSelection()
    {
        var app = new FakeApp();
        var anchor = new FakeRange(2,3,"=GKD(A1,B1)","done");
        var child = new FakeRange(3,3,null,"child") { ParentSpill = anchor };
        var prompt = new FakeRange(8,1,"=GKD.PROMPT(A1)","prompt");
        var unrelated = new FakeRange(9,1,"=SUM(A1:A2)",12.0);
        app.Selection=FakeRange.Selection(new FakeRange(12,4,"=GKD(A1,B1)","scalar"){NullSpillParent=true});
        Invoke("RunCore",app,"aiRefresh",new Func<string,bool>(s=>true));
        Check(GkdAiFunctions.Refreshed.Count==1 && GkdAiFunctions.Refreshed[0]=="7:11:3" && app.StatusBar.ToString().Contains("1 个"),"scalar refresh falls back to cell when SpillParent is null");
        GkdAiFunctions.Refreshed.Clear();
        app.Selection = FakeRange.Selection(child,anchor,child,prompt,unrelated);
        Invoke("RunCore",app,"aiRefresh",new Func<string,bool>(s => true));
        Check(GkdAiFunctions.Refreshed.Count==1 && GkdAiFunctions.Refreshed[0]=="7:1:2", "refresh deduplicated spill anchor stable identity");
        Check(anchor.Dirties==1 && anchor.Calculates==1 && unrelated.Calculates==0 && prompt.Calculates==0,"only GKD anchor recalculated");
        Check(Equals(anchor.Formula,"=GKD(A1,B1)"),"refresh never rewrites formula");
        app.Selection=FakeRange.Selection(); app.Selection.CountLarge=10001;
        bool rejected=false;try{Invoke("RunCore",app,"aiRefresh",new Func<string,bool>(s=>true));}catch(InvalidOperationException){rejected=true;}
        Check(rejected && anchor.Calculates==1,"oversized selection rejected without action");
        app.Selection=null;
        Invoke("RunCore",app,"aiCancel",new Func<string,bool>(s=>true));
        Check(GkdAiFunctions.Scheduler.Cancels==1 && app.StatusBar.ToString().Contains("取消"),"cancel works without range selection and reports status");
        Console.WriteLine("PASS refresh selection, spill deduplication, limits and cancel");
    }
    static void FreezeSelection()
    {
        var app=new FakeApp();
        var anchor=new FakeRange(10,3,"=GKD(A1,B1)","name");
        var spill=new FakeRange(10,3,null,new object[,]{{"name",2.0},{"other",3.0}}){CountLarge=4};
        anchor.Spilling=spill;
        var child=new FakeRange(11,4,null,3.0){ParentSpill=anchor};
        var input=new FakeRange(1,1,"=SUM(A2:A3)",5.0);
        var ordinarySpillAnchor=new FakeRange(20,1,"=SEQUENCE(2)",1.0);
        var ordinarySpillChild=new FakeRange(21,1,null,2.0){ParentSpill=ordinarySpillAnchor};
        var blank=new FakeRange(30,1,null,null);
        app.Selection=FakeRange.Selection(child,anchor,input,ordinarySpillAnchor,ordinarySpillChild,blank);
        string confirmation=null;
        Invoke("RunCore",app,"aiValues",new Func<string,bool>(s=>{confirmation=s;return false;}));
        Check(spill.Pastes==0 && app.EnableEvents,"declining confirmation writes nothing");
        Check(confirmation.Contains("4") && confirmation.Contains(spill.Address) && confirmation.Contains("完整溢出"),"confirmation names full spill extent and count");
        spill.App=app;
        Invoke("RunCore",app,"aiValues",new Func<string,bool>(s=>true));
        Check(spill.Pastes==1 && spill.Copies==1,"whole spill pasted as native values once");
        Check(input.Pastes==0 && input.Copies==0 && Equals(input.Formula,"=SUM(A2:A3)"),"input and unrelated formula untouched");
        Check(ordinarySpillAnchor.Pastes==0&&ordinarySpillAnchor.Copies==0&&ordinarySpillChild.Pastes==0&&ordinarySpillChild.Copies==0&&Equals(ordinarySpillAnchor.Formula,"=SEQUENCE(2)"),"ordinary formula spills and blanks are excluded from freeze plan");
        Check(app.EnableEvents && Equals(app.CutCopyMode,false),"events and clipboard state restored");
        Console.WriteLine("PASS freeze full spill, confirmation, unrelated references and events");
    }
    static void FreezeSingleScalarWithoutSpill()
    {
        foreach (bool events in new[] { true, false })
        {
            var cell = new FakeRange(4, 2, "=GKD(A1,B1)", "中文 ABC") { NullSpillParent = true, NullSpillingRange = true };
            var app = new FakeApp { Selection = FakeRange.Selection(cell), EnableEvents = events };
            cell.App = app;
            Invoke("RunCore", app, "aiValues", new Func<string, bool>(s => s.Contains(cell.Address)));
            Check(cell.Copies == 1 && cell.Pastes == 1 && cell.Formula == null && Equals(cell.Value2, "中文 ABC"), "null spill range freezes original scalar without changing value");
            Check(app.EnableEvents == events && Equals(app.CutCopyMode, false), "scalar freeze restores original events and clipboard");
        }
        Console.WriteLine("PASS scalar freeze with null spill parent and range");
    }
    static void FreezeGuards()
    {
        var cases=new Dictionary<string,Action<FakeRange>>
        {
            {"waiting #N/A",r=>r.Value2=ExcelError.ExcelErrorNA},
            {"COM error variant",r=>r.Value2=new System.Runtime.InteropServices.ErrorWrapper(unchecked((int)0x800A07FA))},
            {"failure",r=>r.Value2="GKD：请求失败"},
            {"canceled",r=>r.Value2="GKD：已取消"},
            {"ASCII marker",r=>r.Value2="GKD: unavailable"},
            {"protected",r=>r.Worksheet.ProtectContents=true},
            {"merged",r=>r.MergeCells=true},
            {"mixed merged",r=>r.MergeCells=null},
            {"legacy array",r=>r.HasArray=true},
            {"overlimit spill",r=>r.CountLarge=10001},
            {"error in spill",r=>{r.CountLarge=2;r.Value2=new object[,]{{"ok",ExcelError.ExcelErrorValue}};}},
            {"merged spill child",r=>{r.CountLarge=2;r.Cells.Add(new FakeRange(r.Row+1,r.Column,null,"x"){MergeCells=true});}},
        };
        foreach(var test in cases)
        {
            var good=new FakeRange(2,1,"=GKD(A1,B1)","ok");
            var bad=new FakeRange(10,1,"=GKD(A1,B1)","ok");
            test.Value(bad);
            var app=new FakeApp{Selection=FakeRange.Selection(good,bad)};
            bool rejected=false,confirmed=false;
            try{Invoke("RunCore",app,"aiValues",new Func<string,bool>(s=>{confirmed=true;return true;}));}catch(InvalidOperationException){rejected=true;}
            Check(rejected && !confirmed && good.Pastes==0 && good.Copies==0,"preflight all before any write: "+test.Key);
        }
        var first=new FakeRange(2,1,"=GKD(A1,B1)","ok"){CountLarge=6000};
        var second=new FakeRange(4,1,"=GKD(A1,B1)","ok"){CountLarge=5000};
        bool refused=false;try{Invoke("RunCore",new FakeApp{Selection=FakeRange.Selection(first,second)},"aiValues",new Func<string,bool>(s=>true));}catch(InvalidOperationException){refused=true;}
        Check(refused && first.Pastes==0,"aggregate full spill limit preflight");
        Console.WriteLine("PASS freeze preflight guards for all targets and aggregate limits");
    }
    static void FreezeFailureRestoresEvents()
    {
        foreach(bool initial in new[]{true,false})
        {
            var cell=new FakeRange(2,1,"=GKD(A1,B1)","ok"){ThrowOnPaste=true};
            var app=new FakeApp{Selection=FakeRange.Selection(cell),EnableEvents=initial};
            cell.App=app;
            bool failed=false;try{Invoke("RunCore",app,"aiValues",new Func<string,bool>(s=>true));}catch(System.Runtime.InteropServices.COMException){failed=true;}
            Check(failed && app.EnableEvents==initial && Equals(app.CutCopyMode,false),"paste failure restores original events state");
        }
        Console.WriteLine("PASS paste failure restores events and clears clipboard mode");
    }
    static void ConfirmationCannotRedirectWrite()
    {
        foreach(bool changeSpill in new[]{false,true})
        {
            var cell=new FakeRange(2,1,"=GKD(A1,B1)","ok");
            var app=new FakeApp{Selection=FakeRange.Selection(cell)};
            bool refused=false;
            try{Invoke("RunCore",app,"aiValues",new Func<string,bool>(s=>{if(changeSpill)cell.Spilling=new FakeRange(2,1,null,new object[,]{{"a","b"}}){CountLarge=2};else cell.Formula="=SUM(A1:B1)";return true;}));}
            catch(InvalidOperationException){refused=true;}
            Check(refused && cell.Copies==0 && cell.Pastes==0 && app.EnableEvents,"changed formula or spill after confirmation rejected before any write");
        }
        Console.WriteLine("PASS confirmation-time formula/spill mutation rejected before writing");
    }
    static void MixedSelectionSkipsNonGkdCells()
    {
        GkdAiFunctions.Refreshed.Clear();
        var app=new FakeApp();
        var ordinary=new FakeRange(2,2,"=SUM(A1:A2)",3.0);
        var blank=new FakeRange(3,2,null,null);
        var ordinarySpill=new FakeRange(4,2,"=SEQUENCE(2)",1.0);
        var spillChild=new FakeRange(5,2,null,2.0){ParentSpill=ordinarySpill};
        var gkd=new FakeRange(6,2,"=GKD(A1,B1)","ok");
        var customAi=new FakeRange(7,2,"=AI_WORKFLOW(A1)","ok");
        var customDeterministic=new FakeRange(8,2,"=DETERMINISTIC_ALIAS(A1)",2.0);
        app.Selection=FakeRange.Selection(ordinary,blank,spillChild,gkd,customAi,customDeterministic);
        Invoke("RunCore",app,"aiRefresh",new Func<string,bool>(s=>true));
        Check(GkdAiFunctions.Refreshed.Count==2&&gkd.Calculates==1&&customAi.Calculates==1&&customDeterministic.Calculates==0&&ordinary.Calculates==0&&ordinarySpill.Calculates==0,"mixed refresh touches standalone GKD and AI aliases, not deterministic custom formulas");

        var onlyOrdinary=new FakeApp{Selection=FakeRange.Selection(ordinary,blank,spillChild)};
        Invoke("RunCore",onlyOrdinary,"aiRefresh",new Func<string,bool>(s=>true));
        Check(onlyOrdinary.StatusBar.ToString().Contains("没有 GKD")&&ordinary.Calculates==0&&ordinarySpill.Calculates==0,"selection with only ordinary formulas and blanks reports status without failing");

        var gkdSpill=new FakeRange(8,2,"=GKD(A1,B1)","ok");
        var gkdSpillRange=new FakeRange(8,2,null,new object[,]{{"ok",1.0}}){CountLarge=2};
        gkdSpill.Spilling=gkdSpillRange;
        var gkdChild=new FakeRange(8,3,null,1.0){ParentSpill=gkdSpill};
        var otherFormula=new FakeRange(9,2,"=SUM(A1:A2)",4.0);
        var freezeBlank=new FakeRange(10,2,null,null);
        var alias=new FakeRange(11,2,"=AI_WORKFLOW(A1)","ok");
        var aliasSpill=new FakeRange(11,2,null,new object[,]{{"ok",2.0}}){CountLarge=2};
        alias.Spilling=aliasSpill;
        var aliasChild=new FakeRange(11,3,null,2.0){ParentSpill=alias};
        var deterministic=new FakeRange(12,2,"=DETERMINISTIC_ALIAS(A1)",2.0);
        var freezeApp=new FakeApp{Selection=FakeRange.Selection(gkdChild,otherFormula,freezeBlank,aliasChild,deterministic)};
        gkdSpillRange.App=freezeApp;
        aliasSpill.App=freezeApp;
        Invoke("RunCore",freezeApp,"aiValues",new Func<string,bool>(s=>true));
        Check(gkdSpillRange.Pastes==1&&gkdSpillRange.Copies==1&&aliasSpill.Pastes==1&&aliasSpill.Copies==1&&otherFormula.Pastes==0&&otherFormula.Copies==0&&deterministic.Pastes==0&&deterministic.Copies==0&&Equals(otherFormula.Formula,"=SUM(A1:A2)"),"mixed freeze converts selected GKD and custom AI spill ranges only");
        Console.WriteLine("PASS mixed selections skip ordinary formulas/spills/blanks for refresh and freeze");
    }
}

namespace ExcelDna.Integration
{
    public enum ExcelError { ExcelErrorNA, ExcelErrorValue }
    public static class ExcelDnaUtil { public static object Application; }
    public class ExcelReference { public IntPtr SheetId { get { return new IntPtr(7); } } public ExcelReference(int r1,int r2,int c1,int c2) {} }
}
namespace System.Windows.Forms
{
    public enum DialogResult { OK, Cancel, Yes, No }
    public enum MessageBoxButtons { OK, YesNo }
    public enum MessageBoxIcon { Information, Warning, Error }
    public static class MessageBox
    {
        public static DialogResult Answer = DialogResult.Yes;
        public static string Last;
        public static DialogResult Show(string message,string title,MessageBoxButtons buttons,MessageBoxIcon icon) { Last=message; return Answer; }
    }
}
namespace Outline.Addin
{
    internal static class GkdCustomFormulaFunctions
    {
        static readonly Dictionary<string,bool> Kinds=new Dictionary<string,bool>(StringComparer.OrdinalIgnoreCase);
        public static bool IsCustomName(string name){return Kinds.ContainsKey(name)||string.Equals(name,"GKD_CUSTOM",StringComparison.OrdinalIgnoreCase);}
        public static bool IsAiFormula(string name){bool ai;return Kinds.TryGetValue(name,out ai)&&ai;}
        public static void SetKind(string name,bool ai){Kinds[name]=ai;}
    }
    internal sealed class GkdAiScheduler { public int Cancels; public void CancelAll() { Cancels++; } }
    internal static class GkdAiFunctions
    {
        internal static readonly GkdAiScheduler Scheduler = new GkdAiScheduler();
        internal static readonly List<string> Refreshed = new List<string>();
        internal static string CellIdentity(IntPtr sheet,int row,int column) { return sheet + ":" + row + ":" + column; }
        internal static void RefreshCell(string id) { Refreshed.Add(id); }
    }
}
