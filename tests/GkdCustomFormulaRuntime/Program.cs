using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Outline.Addin;
class Program
{
    static int checks;
    internal static void Check(bool value, string message) { checks++; if (!value) throw new Exception(message); }
    internal static GkdCustomFormula Formula(string name = "PIPE_MASS") { return new GkdCustomFormula { ExcelName = name, DisplayName = "管道质量", Description = "质量计算", Inputs = new List<string> { "DIAMETER", "LENGTH" }, FormulaExpression = "DIAMETER*LENGTH" }; }
    static Type Required(string name) { var t = typeof(Program).Assembly.GetType("Outline.Addin." + name); Check(t != null, name + " implementation missing"); return t; }
    static void Main()
    {
        var store = Required("GkdCustomFormulaStore");
        var path = Path.Combine(Path.GetTempPath(), "gkd-custom-runtime-" + Guid.NewGuid().ToString("N"), "formulas.json");
        try
        {
            var load = store.GetMethod("Load", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, null, new[] { typeof(string) }, null);
            var save = store.GetMethod("Save", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, null, new[] { typeof(List<GkdCustomFormula>), typeof(string) }, null);
            Check(load != null && save != null, "path overloads for real filesystem testing");
            Check(((List<GkdCustomFormula>)load.Invoke(null, new object[] { path })).Count == 0, "missing store is empty");
            var formula = Formula();
            save.Invoke(null, new object[] { new List<GkdCustomFormula> { formula }, path });
            Check(((List<GkdCustomFormula>)load.Invoke(null, new object[] { path }))[0].ExcelName == "PIPE_MASS", "store round trips validated definitions");
            string before = File.ReadAllText(path);
            formula.FormulaExpression = "Process.Start(1)";
            try { save.Invoke(null, new object[] { new List<GkdCustomFormula> { formula }, path }); throw new Exception("invalid save accepted"); }
            catch (TargetInvocationException) { }
            Check(File.ReadAllText(path) == before, "invalid save preserves previous file byte for byte");
            formula.FormulaExpression = "DIAMETER+LENGTH";
            save.Invoke(null, new object[] { new List<GkdCustomFormula> { formula }, path });
            Check(((List<GkdCustomFormula>)load.Invoke(null, new object[] { path }))[0].FormulaExpression == "DIAMETER+LENGTH", "atomic replacement updates existing file");
            Check(Directory.GetFiles(Path.GetDirectoryName(path), "*.tmp").Length == 0, "save cleans staging files");
            File.WriteAllText(path, "{broken");
            try { load.Invoke(null, new object[] { path }); throw new Exception("corrupt load accepted"); } catch (TargetInvocationException) { }
            Check(File.ReadAllText(path) == "{broken", "load does not overwrite corruption");
        }
        finally { if (Directory.Exists(Path.GetDirectoryName(path))) Directory.Delete(Path.GetDirectoryName(path), true); }
        var runtime = Required("GkdCustomFormulaFunctions");
        var apply = runtime.GetMethod("ApplyDefinitions", BindingFlags.Static | BindingFlags.NonPublic);
        Check(apply != null, "validated runtime snapshot seam");
        var definitions = new List<GkdCustomFormula> { Formula() };
        apply.Invoke(null, new object[] { definitions });
        var custom = runtime.GetMethod("Custom");
        Check(Equals(custom.Invoke(null, new object[] { "PIPE_MASS", new object[] { 3.0, 4.0 } }), 12.0), "universal safe numeric evaluation");
        var functions = ExcelDna.Integration.ExcelIntegration.Functions;
        Check(functions.ContainsKey("PIPE_MASS") && functions.ContainsKey("GKD_CUSTOM"), "real registration entry sends alias and universal delegates");
        var args = new object[32]; for (int i = 0; i < args.Length; i++) args[i] = ExcelDna.Integration.ExcelMissing.Value; args[0] = 3.0; args[1] = 4.0;
        Check(Equals(functions["PIPE_MASS"].DynamicInvoke(args), 12.0), "Excel scalar wrapper packs trailing missing arguments");
        Check(functions["PIPE_MASS"].GetType().GetMethod("Invoke").GetParameters().Length == 32, "RegisterDelegates gets expanded 32-slot wrapper, never raw paramarray");
        definitions[0].FormulaExpression = "DIAMETER+LENGTH";
        apply.Invoke(null, new object[] { definitions });
        Check(Equals(functions["PIPE_MASS"].DynamicInvoke(args), 7.0), "registered alias looks up current definition each call");
        RuntimeCases.Run();
        Console.WriteLine("PASS " + checks + " custom formula store/runtime checks");
    }
}
