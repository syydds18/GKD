using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

class Program
{
    static Type Formatter { get { return typeof(Program).Assembly.GetType("Outline.Addin.GkdAiTableFormatting"); } }
    public static void Check(bool ok,string message) { if(!ok) throw new Exception(message); }
    static object Call(string method,params object[] args)
    {
        Check(Formatter!=null,"deferred GKD formatter is missing");
        try { return Formatter.GetMethod(method,BindingFlags.Static|BindingFlags.NonPublic|BindingFlags.Public).Invoke(null,args); }
        catch(TargetInvocationException e) { throw e.InnerException; }
    }
    public static object[,] Table() { return new object[,] {{"结果", ""},{"", ""},{"项目", "金额"},{"中文",12.5}}; }
    static void Main()
    {
        var sheet=new FakeSheet(); var data=Table(); var anchor=sheet.Load(data);
        Check(Call("TryFormatTable",anchor,anchor.Formula2,data).ToString()=="Applied","ready owned spill formatted");
        Check(sheet.Log.Contains("1:1:1:2|Align|7"),"center across selection on title");
        Check(sheet.Log.Contains("1:1:4:2|Font.Size|10.5"),"uniform table font size");
        Check(sheet.Log.Contains("1:1:4:2|Font.Name|宋体"),"uniform table Chinese font");
        Check(sheet.Log.IndexOf("1:1:4:2|Font.Name|宋体") < sheet.Log.IndexOf("1:1:4:2|Font.Name|Times New Roman"),"uniform font order Song then Times New Roman");
        Check(sheet.Log.Contains("1:1:2:2|Borders.LineStyle|-4142"),"top two rows no borders");
        Check(sheet.Log.Any(x=>x.StartsWith("1:1:4:2|Font.Name|Times New Roman")),"table fonts finish at Times New Roman");
        Check(sheet.Log.Contains("3:1:2:2|Borders.LineStyle|1") && sheet.Log.Contains("3:1:2:2|Borders.Weight|2") && sheet.Log.Contains("3:1:2:2|Borders.Color|0"),"header and data continuous thin black borders");
        Check(sheet.Log.Contains("1:1:4:2|Font.Name|Times New Roman"),"table fonts finish at Times New Roman");
        Check(sheet.Log.Contains("3:1:2:2|Borders[5].LineStyle|-4142") && sheet.Log.Contains("3:1:2:2|Borders[6].LineStyle|-4142"),"no diagonal borders in grid");
        Check(!sheet.Log.Any(x=>x.Contains("Merge")||x.Contains("Value")||x.Contains("Formula")||x.Contains("Interior")),"no values, formulas, merges or fill writes");
        Console.WriteLine("PASS title, spacer, body, border extent and non-destructive writes");
        Ownership();
        Scalars();
        Scheduling();
    }
    static void Scheduling()
    {
        var sheet=new FakeSheet();var a=sheet.Load(Table());var app=new FakeApp(sheet);
        ExcelDna.Integration.ExcelDnaUtil.Application=app;
        string formula=a.Formula2;
        ExcelDna.Integration.XlCall.Handler=(id,args)=>id==ExcelDna.Integration.XlCall.xlfGetCell?(object)formula:"[Book.xlsx]Target";
        var caller=new ExcelDna.Integration.ExcelReference {SheetId=new IntPtr(713),RowFirst=0,ColumnFirst=0};
        var table=Table();Call("Schedule",caller,table);
        Check(ExcelDna.Integration.ExcelAsyncUtil.Queue.Count==1 && sheet.Log.Count==0 && app.Accesses==0,"UDF scheduling never accesses COM");
        table[0,0]="mutated";
        ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check(sheet.Log.Count>0 && app.Accesses==1,"deferred macro resolves exact workbook and uses copied array");
        sheet.Log.Clear();Call("Schedule",caller,Table());
        Check(ExcelDna.Integration.ExcelAsyncUtil.Queue.Count==0,"unchanged final result deduplicated");
        caller=new ExcelDna.Integration.ExcelReference {SheetId=new IntPtr(714),RowFirst=0,ColumnFirst=0};
        Call("Schedule",caller,Table());formula="=GKD(A9,B9,\"表格\")";a.Formula2=formula;
        ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check(sheet.Log.Count==0,"formula changed after scheduling aborts callback");
        caller=new ExcelDna.Integration.ExcelReference {SheetId=new IntPtr(715),RowFirst=0,ColumnFirst=0};
        a.Spill=null;Call("Schedule",caller,Table());ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check(System.Windows.Forms.Timer.Active.Count==1,"pending spill retries on UI timer");
        for(int i=0;i<30;i++)
        {
            foreach(var timer in System.Windows.Forms.Timer.Active.ToArray())timer.Fire();
            while(ExcelDna.Integration.ExcelAsyncUtil.Queue.Count>0)ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        }
        Check(sheet.Log.Count==0 && System.Windows.Forms.Timer.Active.Count==0,"retry budget exhausted without touching ordinary data");
        caller=new ExcelDna.Integration.ExcelReference {SheetId=new IntPtr(720),RowFirst=0,ColumnFirst=0};
        a.Spill=new FakeRange(sheet,1,1,3,2);Call("Schedule",caller,Table());ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check(System.Windows.Forms.Timer.Active.Count==1,"old spill dimensions retry while final spill is publishing");
        a.Spill=new FakeRange(sheet,1,1,4,2);
        foreach(var timer in System.Windows.Forms.Timer.Active.ToArray())timer.Fire();
        while(ExcelDna.Integration.ExcelAsyncUtil.Queue.Count>0)ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check(sheet.Log.Count>0 && System.Windows.Forms.Timer.Active.Count==0,"ready spill formatted on a later macro");sheet.Log.Clear();
        caller=new ExcelDna.Integration.ExcelReference {SheetId=new IntPtr(716),RowFirst=0,ColumnFirst=0};
        a.Spill=null;sheet.Values=new object[,]{{"scalar ABC"}};Call("ScheduleScalar",caller,"scalar ABC");
        ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check(sheet.Log.SequenceEqual(new[]{"1:1:1:1|Font.Name|宋体","1:1:1:1|Font.Name|Times New Roman","1:1:1:1|Font.Size|10.5"}),"scalar public scheduling path");
        sheet.Log.Clear();Call("ScheduleScalar",caller,"GKD：请求失败");
        Check(ExcelDna.Integration.ExcelAsyncUtil.Queue.Count==0 && sheet.Log.Count==0,"scalar diagnostic excluded before scheduling");
        Console.WriteLine("PASS macro-only COM, workbook resolution, copied snapshots, deduplication, stale callbacks and bounded UI retries");
        caller=new ExcelDna.Integration.ExcelReference {SheetId=new IntPtr(717),RowFirst=0,ColumnFirst=0};
        Call("ScheduleScalar",caller,"different");ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check(System.Windows.Forms.Timer.Active.Count==1,"pending scalar timer present before unload");
        Call("Stop");Check(System.Windows.Forms.Timer.Active.Count==0,"unload disposes timers");
        Call("ScheduleScalar",caller,"scalar ABC");Check(ExcelDna.Integration.ExcelAsyncUtil.Queue.Count==0,"unloaded formatter ignores scheduling");
        Call("Start");Call("ScheduleScalar",caller,"scalar ABC");
        Call("Stop");ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();Check(sheet.Log.Count==0,"queued callbacks inert after unload");
        Call("Start");
        sheet.Values[0,0]="中文ABC123";app.StatusBar="prior status";
        caller=new ExcelDna.Integration.ExcelReference {SheetId=new IntPtr(718),RowFirst=0,ColumnFirst=0};
        Call("ScheduleScalar",caller,"中文ABC123");ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check(Convert.ToString(app.StatusBar).Contains("混排") && System.Windows.Forms.Timer.Active.Count==1,"mixed-script font limitation shown nonmodally");
        app.StatusBar="user status";
        foreach(var timer in System.Windows.Forms.Timer.Active.ToArray())timer.Fire();
        while(ExcelDna.Integration.ExcelAsyncUtil.Queue.Count>0)ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check((string)app.StatusBar=="user status" && System.Windows.Forms.Timer.Active.Count==0,"notice expiry preserves a newer user status");
        sheet.Log.Clear();caller=new ExcelDna.Integration.ExcelReference {SheetId=new IntPtr(719),RowFirst=0,ColumnFirst=0};
        Call("ScheduleScalar",caller,"中文ABC123");ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        foreach(var timer in System.Windows.Forms.Timer.Active.ToArray())timer.Fire();
        while(ExcelDna.Integration.ExcelAsyncUtil.Queue.Count>0)ExcelDna.Integration.ExcelAsyncUtil.Queue.Dequeue()();
        Check((string)app.StatusBar=="user status","notice expiry restores captured status when still owned");
        Call("Stop");
        Console.WriteLine("PASS unload lifecycle and temporary mixed-script limitation notice");
    }
    static void Ownership()
    {
        var sheet=new FakeSheet(); var a=sheet.Load(Table()); string f=a.Formula2;
        a.Formula2="=SUM(A1:A2)";
        Check(Call("TryFormatTable",a,f,Table()).ToString()=="Rejected" && sheet.Log.Count==0,"stale formula rejected before any formatting");
        a.Formula2=f; a.Spill=null;
        Check(Call("TryFormatTable",a,f,Table()).ToString()=="NotReady" && sheet.Log.Count==0,"ordinary data never formatted as a spill");
        a.Spill=new FakeRange(sheet,1,1,3,2);
        Check(Call("TryFormatTable",a,f,Table()).ToString()=="NotReady" && sheet.Log.Count==0,"wrong spill dimensions never formatted");
        a.Spill=new FakeRange(sheet,2,1,4,2);
        Check(Call("TryFormatTable",a,f,Table()).ToString()=="Rejected" && sheet.Log.Count==0,"wrong spill anchor rejected");
        a.Spill=new FakeRange(sheet,1,1,4,2);sheet.Values[3,1]=99.0;
        Check(Call("TryFormatTable",a,f,Table()).ToString()=="NotReady" && sheet.Log.Count==0,"all spill values checked, not title alone");
        sheet.Values=Table();a.Spill.MergeCells=true;
        Check(Call("TryFormatTable",a,f,Table()).ToString()=="Rejected" && sheet.Log.Count==0,"existing merge rejected without unmerging");
        var bad=Table();bad[1,1]="not blank";
        Check(!(bool)Call("ValidLayout",bad),"spacer layout mandatory");bad=Table();bad[0,1]="unexpected";
        Check(!(bool)Call("ValidLayout",bad),"only first title cell populated");
        var oversized=new object[10001,1];oversized[0,0]="title";oversized[1,0]="";oversized[2,0]="header";
        Check(!(bool)Call("ValidLayout",oversized),"same 10000-cell output cap as protocol");
        foreach(string formula in new[]{"=GKD(A1,B1)+1","=GKD.PROMPT(A1)","=IF(TRUE,GKD(A1,B1),0)","=GKD(A1"})
            Check(!(bool)Call("IsGkdFormula",formula),"unrelated formula rejected: "+formula);
        Check((bool)Call("IsGkdFormula"," =@_xll.gkd(\"x )\"\"\",'Sheet''1'!B1)"),"quoted formula accepted");
        Console.WriteLine("PASS stale ownership, whole-spill verification, dimensions, merge refusal and mandatory layout");
    }
    static void Scalars()
    {
        var sheet=new FakeSheet();var a=sheet.Load(new object[,]{{"中文 ABC123"}});a.Spill=null;string formula=a.Formula2;
        Check(Call("TryFormatScalar",a,formula,"中文 ABC123").ToString()=="Applied","owned scalar formatted");
        Check(sheet.Log.SequenceEqual(new[]{"1:1:1:1|Font.Name|宋体","1:1:1:1|Font.Name|Times New Roman","1:1:1:1|Font.Size|10.5"}),"scalar only applies requested Chinese-to-Western font sequence and size");
        sheet.Log.Clear();sheet.Values[0,0]="ABC 123";
        Check(Call("TryFormatScalar",a,formula,"ABC 123").ToString()=="Applied" && sheet.Log.Contains("1:1:1:1|Font.Name|Times New Roman"),"Latin scalar Times New Roman");
        sheet.Log.Clear();a.Spill=a;
        Check(Call("TryFormatScalar",a,formula,"ABC 123").ToString()=="Rejected" && sheet.Log.Count==0,"scalar refuses spills");
        a.Spill=null;
        Check(Call("TryFormatScalar",a,formula,"obsolete").ToString()=="NotReady" && sheet.Log.Count==0,"stale scalar result not styled");
        Check(Call("TryFormatScalar",a,formula,"GKD：已取消").ToString()=="Rejected","diagnostic strings never styled");
        a.SpillError=unchecked((int)0x80020003);
        Check(Call("TryFormatScalar",a,formula,"ABC 123").ToString()=="Applied","legacy unsupported HasSpill permits verified scalar formatting");
        a.SpillError=unchecked((int)0x800AC472);sheet.Log.Clear();bool busy=false;
        try{Call("TryFormatScalar",a,formula,"ABC 123");}catch(System.Runtime.InteropServices.COMException){busy=true;}
        Check(busy && sheet.Log.Count==0,"busy COM errors are not mistaken for unsupported spill API");
        Console.WriteLine("PASS scalar-only font, result guards and explicit mixed-script fallback");
    }
}
public class FakeSheet
{
    public readonly List<string> Log=new List<string>(); public object[,] Values; public FakeRange Anchor;
    public FakeCells Cells { get { return new FakeCells(this,1,1); } }
    public FakeRanges Range { get { return new FakeRanges(this); } }
    public FakeRange Load(object[,] values) { Values=values; Anchor=new FakeRange(this,1,1,1,1); Anchor.Formula2="=GKD(A1,B1,\"表格\")"; Anchor.Spill=new FakeRange(this,1,1,values.GetLength(0),values.GetLength(1)); return Anchor; }
}
public class FakeCells
{
    readonly FakeSheet s; readonly int row,col;
    public FakeCells(FakeSheet sheet,int r,int c) { s=sheet;row=r;col=c; }
    public FakeRange this[int r,int c] { get { if(row+r-1==1 && col+c-1==1 && s.Anchor!=null)return s.Anchor; return new FakeRange(s,row+r-1,col+c-1,1,1); } }
}
public class FakeRanges
{
    readonly FakeSheet s; public FakeRanges(FakeSheet sheet) {s=sheet;}
    public FakeRange this[FakeRange first,FakeRange last] { get { return new FakeRange(s,first.Row,first.Column,last.Row-first.Row+1,last.Column-first.Column+1); } }
}
public class FakeCount { public int Count {get;set;} }
public class FakeRange
{
    public FakeSheet Worksheet {get;private set;} public int Row {get;private set;} public int Column {get;private set;}
    public FakeCount Rows {get;private set;} public FakeCount Columns {get;private set;}
    public string Formula2 {get;set;} public bool HasFormula {get {return !string.IsNullOrEmpty(Formula2);} }
    public FakeRange Spill; public int? SpillError; public bool HasSpill {get {if(SpillError.HasValue)throw new System.Runtime.InteropServices.COMException("HasSpill unavailable",SpillError.Value);return Spill!=null;} }
    public FakeRange SpillingToRange {get { if(Spill==null) throw new InvalidOperationException("no spill"); return Spill; } }
    public bool MergeCells {get;set;}
    public FakeRange SpillParent { get { return Worksheet.Anchor; } }
    public FakeCells Cells {get {return new FakeCells(Worksheet,Row,Column);} }
    public FakeFont Font {get {return new FakeFont(this);} }
    public FakeCharacters Characters(int start,int length){return new FakeCharacters(this,start,length);}
    public FakeBorders Borders {get {return new FakeBorders(this);} }
    public object HorizontalAlignment {set {Record("Align",value);} }
    public string Key {get {return Row+":"+Column+":"+Rows.Count+":"+Columns.Count;} }
    public FakeRange(FakeSheet s,int r,int c,int nr,int nc) {Worksheet=s;Row=r;Column=c;Rows=new FakeCount{Count=nr};Columns=new FakeCount{Count=nc};}
    public void Record(string property,object value) {Worksheet.Log.Add(Key+"|"+property+"|"+Convert.ToString(value,System.Globalization.CultureInfo.InvariantCulture));}
    public object Value2
    {
        get
        {
            if(Rows.Count==1 && Columns.Count==1) return Worksheet.Values[Row-1,Column-1];
            var a=Array.CreateInstance(typeof(object),new[]{Rows.Count,Columns.Count},new[]{1,1});
            for(int r=1;r<=Rows.Count;r++)for(int c=1;c<=Columns.Count;c++)a.SetValue(Worksheet.Values[Row+r-2,Column+c-2],r,c);
            return a;
        }
    }
}
public class FakeCharacters
{
    readonly FakeRange range;readonly int start,length;public FakeCharacters(FakeRange r,int s,int l){range=r;start=s;length=l;}
    public FakeCharacterFont Font {get {return new FakeCharacterFont(range,start,length);} }
}
public class FakeCharacterFont
{
    readonly FakeRange range;readonly int start,length;public FakeCharacterFont(FakeRange r,int s,int l){range=r;start=s;length=l;}
    public string Name {set{range.Record("Characters["+start+","+length+"].Font.Name",value);} }
}
public class FakeFont
{
    readonly FakeRange r; public FakeFont(FakeRange range){r=range;}
    public string Name {set{r.Record("Font.Name",value);} } public double Size {set{r.Record("Font.Size",value);} } public bool Bold {set{r.Record("Font.Bold",value);} }
}
public class FakeBorders
{
    readonly FakeRange r; public FakeBorders(FakeRange range){r=range;}
    public FakeBorders this[int index] {get {return new FakeBorders(r,index);} }
    readonly int? index;
    public FakeBorders(FakeRange range,int i){r=range;index=i;}
    public int LineStyle {set{r.Record(index.HasValue?"Borders["+index.Value+"].LineStyle":"Borders.LineStyle",value);} } public int Weight {set{r.Record("Borders.Weight",value);} } public int Color {set{r.Record("Borders.Color",value);} }
}
public class FakeApp
{
    readonly FakeSheet sheet; public int Accesses; public object StatusBar;
    public FakeApp(FakeSheet s){sheet=s;}
    public FakeBooks Workbooks {get {Accesses++;return new FakeBooks(sheet);} }
}
public class FakeBooks
{
    readonly FakeSheet sheet;public FakeBooks(FakeSheet s){sheet=s;}
    public FakeBook this[string name] {get {Program.Check(name=="Book.xlsx","explicit workbook lookup");return new FakeBook(sheet);} }
}
public class FakeBook
{
    readonly FakeSheet sheet;public FakeBook(FakeSheet s){sheet=s;}
    public FakeSheets Worksheets {get {return new FakeSheets(sheet);} }
}
public class FakeSheets
{
    readonly FakeSheet sheet;public FakeSheets(FakeSheet s){sheet=s;}
    public FakeSheet this[string name] {get {Program.Check(name=="Target","explicit sheet lookup");return sheet;} }
}
namespace ExcelDna.Integration
{
    public class ExcelReference { public IntPtr SheetId; public int RowFirst,ColumnFirst; }
    public static class ExcelDnaUtil { public static object Application; }
    public static class XlCall { public const int xlSheetNm=1,xlfGetCell=2; public static Func<int,object[],object> Handler; public static object Excel(int function,params object[] args){return Handler(function,args);} }
    public static class ExcelAsyncUtil { public static readonly Queue<Action> Queue=new Queue<Action>(); public static void QueueAsMacro(Action action){Queue.Enqueue(action);} }
}
namespace System.Windows.Forms
{
    public class Timer : IDisposable { public static readonly List<Timer> Active=new List<Timer>(); public int Interval {get;set;} public event EventHandler Tick; public void Start(){Active.Add(this);} public void Stop(){Active.Remove(this);} public void Dispose(){Stop();} public void Fire(){if(Tick!=null)Tick(this,EventArgs.Empty);} }
}
