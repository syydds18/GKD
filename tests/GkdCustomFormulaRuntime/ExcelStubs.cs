using System;
using System.Collections.Generic;
using ExcelDna.Integration;
namespace ExcelDna.Integration
{
    [AttributeUsage(AttributeTargets.Method)] public class ExcelFunctionAttribute : Attribute { public string Name, Description, Category; public bool IsMacroType, IsThreadSafe, ExplicitRegistration, SuppressOverwriteError; }
    [AttributeUsage(AttributeTargets.Method)] public class ExcelCommandAttribute : Attribute { public string Name; }
    [AttributeUsage(AttributeTargets.Parameter)] public class ExcelArgumentAttribute : Attribute { public string Name, Description; public bool AllowReference; }
    public enum ExcelError { ExcelErrorNA, ExcelErrorValue, ExcelErrorDiv0 }
    public class ExcelMissing { public static readonly ExcelMissing Value = new ExcelMissing(); }
    public class ExcelEmpty { public static readonly ExcelEmpty Value = new ExcelEmpty(); }
    public sealed class ExcelReference { public object[,] Values; public int Reads; public bool Oversized; public int RowFirst, RowLast, ColumnFirst, ColumnLast; public List<ExcelReference> InnerReferences = new List<ExcelReference>(); }
    public static class ExcelDnaUtil { public static bool InWizard; public static bool IsInFunctionWizard() { return InWizard; } }
    public static class XlCall { public const int xlfSum = 4, xlfIf = 1, xlfConvert = 468; }
    public static class ExcelIntegration
    {
        public static readonly Dictionary<string, Delegate> Functions = new Dictionary<string, Delegate>();
        public static readonly Dictionary<string, ExcelFunctionAttribute> Metadata = new Dictionary<string, ExcelFunctionAttribute>();
        public static readonly Dictionary<string, List<object>> Arguments = new Dictionary<string, List<object>>();
        public static int Calls;
        public static void RegisterDelegates(List<Delegate> functions, List<object> attributes, List<List<object>> arguments)
        {
            Calls++;
            for (int i = 0; i < functions.Count; i++) { var a = (ExcelFunctionAttribute)attributes[i]; Functions[a.Name] = functions[i]; Metadata[a.Name] = a; Arguments[a.Name] = arguments[i]; }
        }
    }
}
namespace Outline.Addin
{
    public static class GkdAiFunctions
    {
        public static int Calls; public static string Prompt; public static object[,] Matrix;
        public static object Gkd(object prompt, object data = null, object outputMode = null) { Calls++; Prompt = (string)prompt; Matrix = (object[,])data; return "AI reply"; }
        internal static object[,] ReadReference(ExcelReference reference) { if (reference.Oversized) throw new ArgumentException("引用过大"); reference.Reads++; return reference.Values; }
    }
}
