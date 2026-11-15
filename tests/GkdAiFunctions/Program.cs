using System;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using System.Collections.Generic;
using Outline.Addin;
using ExcelDna.Integration;
class Program
{
 static void Check(bool b,string m){if(!b)throw new Exception(m);}
 static object Call(object data,object mode){var type=typeof(Program).Assembly.GetType("Outline.Addin.GkdAiFunctions");Check(type!=null,"GKD formula entry missing");return type.GetMethod("Gkd").Invoke(null,new object[]{"总结",data,mode});}
 static object[,] Rows(int rows, string prefix, int columns=1)
 {
  var values=new object[rows,columns];
  for(int r=0;r<rows;r++)for(int c=0;c<columns;c++)values[r,c]=prefix+(r*columns+c);
  return values;
 }
 static void EmptyDataIsLocal()
 {
  var emptyReference=new ExcelReference {RowFirst=0,RowLast=1,ColumnFirst=0,ColumnLast=1,Value=new object[,]{{null,ExcelEmpty.Value},{ExcelMissing.Value," \t\r\n"}}};
  var emptyDisjoint=new ExcelReference();
  emptyDisjoint.InnerReferences.Add(new ExcelReference {RowFirst=0,RowLast=0,ColumnFirst=0,ColumnLast=0,Value=ExcelEmpty.Value});
  emptyDisjoint.InnerReferences.Add(new ExcelReference {RowFirst=2,RowLast=2,ColumnFirst=2,ColumnLast=3,Value=new object[,]{{null," "}}});
  var bounded=(object[,])Array.CreateInstance(typeof(object),new[]{1,2},new[]{1,1});
  bounded[1,1]=ExcelMissing.Value;bounded[1,2]=" ";
  object[] emptyInputs={""," \t\r\n",ExcelEmpty.Value,new object[,]{{null}},new object[,]{{ExcelMissing.Value}},new object[,]{{null,ExcelEmpty.Value},{ExcelMissing.Value," "}},new object[0,0],emptyReference,emptyDisjoint,bounded};
  int profiles=GkdAiStore.Resolves,asyncCalls=ExcelAsyncUtil.Calls,scheduled=GkdAiScheduler.Calls,backend=GkdAiClient.Calls,formats=GkdAiTableFormatting.Scalars+GkdAiTableFormatting.Tables;
  GkdAiStore.ThrowOnResolve=true;
  try
  {
   foreach(object input in emptyInputs)
    foreach(object mode in new object[]{ExcelMissing.Value,1.0})
    {
     Check(Equals(Call(input,mode),"GKD：对应数据为空，请检查。"),"explicit empty scalar, matrix or reference must return exact local message");
     Check(GkdAiStore.Resolves==profiles&&ExcelAsyncUtil.Calls==asyncCalls&&GkdAiScheduler.Calls==scheduled&&GkdAiClient.Calls==backend&&GkdAiTableFormatting.Scalars+GkdAiTableFormatting.Tables==formats,"empty input must not resolve profile, schedule, call backend or format");
    }
   Check(Equals(Call(new object[,]{{null,ExcelError.ExcelErrorValue},{" ",ExcelMissing.Value}},ExcelMissing.Value),ExcelError.ExcelErrorValue),"Excel errors in otherwise empty data propagate before profile lookup");
   XlCall.SupportsArrays=false;
   try {Check(Equals(Call(emptyReference,1.0),"GKD：对应数据为空，请检查。"),"empty table input remains a local empty-data result on legacy Excel");}
   finally {XlCall.SupportsArrays=true;}
  }
  finally {GkdAiStore.ThrowOnResolve=false;}
  Check(Equals(Call(0.0,ExcelMissing.Value),"reply")&&Equals(GkdAiClient.LastData[0,0],0.0),"numeric zero is real scalar data");
  Check(Equals(Call(new object[,]{{null," "},{0.0,ExcelEmpty.Value}},ExcelMissing.Value),"reply")&&Equals(GkdAiClient.LastData[1,0],0.0),"zero prevents mixed matrix from being treated as empty");
  Check(Equals(Call(null,ExcelMissing.Value),"reply")&&GkdAiClient.LastData==null,"optional null default preserves text-only mode");
  Check(Equals(GkdAiFunctions.Gkd("总结"),"reply")&&GkdAiClient.LastData==null,"omitted data preserves pure-text mode");
  Check(Equals(Call(ExcelMissing.Value,ExcelMissing.Value),"reply")&&GkdAiClient.LastData==null,"Excel missing argument preserves pure-text mode");
  Console.WriteLine("PASS empty-data local guard before profile, async scheduling and backend; errors, zero and omitted data preserved");
 }
 static void Main()
 {
  var udf=typeof(GkdAiFunctions).GetMethod("Gkd");
  Check(udf.GetParameters()[0].ParameterType==typeof(object),"prompt must accept ExcelMissing for no-argument welcome");
  Check(udf.Invoke(null,new object[]{ExcelMissing.Value,ExcelMissing.Value,ExcelMissing.Value}).Equals("感谢使用！")&&GkdAiClient.Calls==0,"empty call welcome must work without model and without network");
  Check(udf.Invoke(null,new object[]{ExcelMissing.Value,null,null}).Equals("感谢使用！")&&GkdAiClient.Calls==0,"optional null defaults preserve empty-call welcome without network");
  var disjoint=new ExcelReference();disjoint.InnerReferences.Add(new ExcelReference {RowFirst=0,RowLast=0,ColumnFirst=0,ColumnLast=0,Value="A"});disjoint.InnerReferences.Add(new ExcelReference {RowFirst=2,RowLast=3,ColumnFirst=2,ColumnLast=2,Value=new object[,]{{"B"},{"C"}}});
  var matrix=(object[,])typeof(GkdAiFunctions).GetMethod("ReadReference",BindingFlags.NonPublic|BindingFlags.Static).Invoke(null,new object[]{disjoint});
  Check(matrix.GetLength(0)==3&&matrix.GetLength(1)==1&&Equals(matrix[0,0],"A")&&Equals(matrix[1,0],"B")&&Equals(matrix[2,0],"C"),"explicit disjoint references concatenate in InnerReferences order");
  var sparseWide=new ExcelReference();
  sparseWide.InnerReferences.Add(new ExcelReference {RowFirst=0,RowLast=99,ColumnFirst=0,ColumnLast=0,Value=Rows(100,"R")});
  sparseWide.InnerReferences.Add(new ExcelReference {RowFirst=0,RowLast=0,ColumnFirst=2,ColumnLast=101,Value=Rows(1,"W",100)});
  var sparseMatrix=(object[,])typeof(GkdAiFunctions).GetMethod("ReadReference",BindingFlags.NonPublic|BindingFlags.Static).Invoke(null,new object[]{sparseWide});
  Check(sparseMatrix.GetLength(0)==200&&sparseMatrix.GetLength(1)==1&&Equals(sparseMatrix[0,0],"R0")&&Equals(sparseMatrix[199,0],"W99"),"sparse wide references count actual cells and flatten padded shape");
  Check(Call("x",1.0).Equals("reply")&&GkdAiClient.Mode=="表格","numeric 1 selects table output");
  Check(Call("x",0.0).Equals("reply")&&GkdAiClient.Mode=="单值","numeric 0 selects scalar output");
  Check(Call(ExcelMissing.Value,1.0).Equals("reply")&&GkdAiClient.LastData==null&&GkdAiClient.Mode=="表格","text-only table request uses omitted data and explicit table mode");
  GkdAiTableFormatting.Scalars=0;GkdAiClient.Calls=0;
  ExcelDnaUtil.InWizard=true;Check(Call("secret",ExcelMissing.Value).Equals(ExcelError.ExcelErrorNA)&&GkdAiClient.Calls==0,"formula wizard must not issue requests");ExcelDnaUtil.InWizard=false;
  Check(Call(ExcelMissing.Value,ExcelMissing.Value).Equals("reply")&&GkdAiClient.LastData==null,"text-only prompt does not fabricate a data matrix");
  Check(Call(ExcelError.ExcelErrorValue,ExcelMissing.Value).Equals(ExcelError.ExcelErrorValue),"source errors propagate without request");
  Check(Call("x","unknown").ToString().Contains("输出"),"unknown output mode rejected");
  Check(Call("x",ExcelMissing.Value).Equals("reply"),"scalar request result");
  Check(GkdAiClient.LastData.GetLength(0)==1&&Equals(GkdAiClient.LastData[0,0],"x"),"single cell becomes explicit matrix");
  Check(GkdAiTableFormatting.Scalars==2,"scalar typography queued outside background work");
  ExcelAsyncUtil.Pending=true;Check(Call("x",ExcelMissing.Value).Equals(ExcelError.ExcelErrorNA)&&GkdAiTableFormatting.Scalars==2,"pending result is not formatted");ExcelAsyncUtil.Pending=false;
  XlCall.SupportsArrays=false;Check(Call("x","表格").ToString().Contains("动态数组"),"legacy Excel must not silently truncate a generated table");XlCall.SupportsArrays=true;
  Call(new object[,]{{"名称","数量"},{"材料",2.0}},"表格");Check(GkdAiClient.Mode=="表格"&&GkdAiClient.LastData.GetLength(1)==2,"2D input preserved and table mode forwarded");
  var type=typeof(Program).Assembly.GetType("Outline.Addin.GkdAiFunctions");
  type.GetMethod("RefreshCell",BindingFlags.NonPublic|BindingFlags.Static).Invoke(null,new object[]{"one"});
  type.GetMethod("RefreshCell",BindingFlags.NonPublic|BindingFlags.Static).Invoke(null,new object[]{"two"});
  var revisions=(System.Collections.Generic.Dictionary<string,long>)type.GetField("Revisions",BindingFlags.NonPublic|BindingFlags.Static).GetValue(null);
  Check(revisions["one"]!=revisions["two"],"separate explicit refreshes need unique request generations to avoid sharing stale refreshed cache");
  EmptyDataIsLocal();
  Console.WriteLine("PASS GKD formula guards, explicit references, scalar and table dispatch (portable model)");
 }
}
namespace ExcelDna.Integration
{
 [AttributeUsage(AttributeTargets.Method)]public class ExcelFunctionAttribute:Attribute{public string Name{get;set;}public string Description{get;set;}public string Category{get;set;}public bool IsThreadSafe{get;set;}public bool IsMacroType{get;set;}}
 [AttributeUsage(AttributeTargets.Parameter)]public class ExcelArgumentAttribute:Attribute{public string Name{get;set;}public string Description{get;set;}public bool AllowReference{get;set;}}
 public enum ExcelError{ExcelErrorNA,ExcelErrorValue}
 public class ExcelMissing{public static readonly object Value=new ExcelMissing();}
 public class ExcelEmpty{public static readonly object Value=new ExcelEmpty();}
 public class ExcelReference{public IntPtr SheetId=new IntPtr(1);public int RowFirst=1,RowLast,ColumnFirst=1,ColumnLast;public List<ExcelReference> InnerReferences=new List<ExcelReference>();public object Value;public object GetValue(){return Value;}}
 public static class XlCall{public const int xlfCaller=89;public static bool SupportsArrays=true;public static object Excel(int n,params object[] args){if(n==614&&!SupportsArrays)throw new Exception("legacy");return n==614?(object)true:new ExcelReference();}}
 public static class ExcelDnaUtil{public static bool InWizard;public static bool IsInFunctionWizard(){return InWizard;}}
 public static class ExcelAsyncUtil{public static bool Pending;public static int Calls;public static object RunTaskWithCancellation<T>(string n,object p,Func<CancellationToken,Task<T>> work){Calls++;return Pending?(object)ExcelError.ExcelErrorNA:work(CancellationToken.None).GetAwaiter().GetResult();}}
}
namespace Outline.Addin
{
 internal static class GkdAiTableFormatting{public static int Scalars,Tables;public static void Stop(){}public static void Start(){}public static void ScheduleScalar(ExcelReference r,string s){Scalars++;}public static void Schedule(ExcelReference r,object[,] d){Tables++;}}
 internal class GkdAiProfile{public void Validate(){}}
 internal static class GkdAiStore{public static int Resolves;public static bool ThrowOnResolve;public static GkdAiProfile ResolveDefault(){Resolves++;if(ThrowOnResolve)throw new ArgumentException("profile lookup must not happen");return new GkdAiProfile();}}
 internal static class GkdAiProtocol{public static string NormalizeMode(string s){if(s=="1")return "表格";if(s=="0")return "单值";if(s=="单值"||s=="表格")return s;throw new ArgumentException("输出模式无效");}public static string BuildPrompt(string p,object[,] d,string m){return p+m;}}
 internal class GkdAiReply{public object Value="reply";}
 internal static class GkdAiClient{public static int Calls;public static object[,] LastData;public static string Mode;public static Task<GkdAiReply> ExecuteAsync(GkdAiProfile p,string q,object[,] d,string m,CancellationToken t){Calls++;LastData=d;Mode=m;return Task.FromResult(new GkdAiReply());}}
 internal class GkdAiScheduler{public static int Calls;public Task<object> RunAsync(string k,Func<CancellationToken,Task<object>> w,CancellationToken t){Calls++;return w(t);}public void CancelAll(){}public void ClearCache(){}}
}
