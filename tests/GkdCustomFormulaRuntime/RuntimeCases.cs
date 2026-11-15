using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Outline.Addin;
static class RuntimeCases
{
    static void Check(bool b, string m) { Program.Check(b, m); }
    static void Reject(Action action, string message) { bool rejected = false; try { action(); } catch (ArgumentException) { rejected = true; } Check(rejected, message); }
    static object Call(params object[] values) { return GkdCustomFormulaFunctions.Custom("PIPE_MASS", values); }
    internal static void Run()
    {
        var customAttribute = (ExcelFunctionAttribute)Attribute.GetCustomAttribute(typeof(GkdCustomFormulaFunctions).GetMethod("Custom"), typeof(ExcelFunctionAttribute));
        Check(customAttribute != null && customAttribute.ExplicitRegistration, "implicit .dna exports must not auto-register the raw paramarray method");
        var f = Program.Formula(); GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { f });
        var alias = ExcelIntegration.Functions["PIPE_MASS"];
        Check(ExcelIntegration.Metadata["PIPE_MASS"].IsMacroType && !ExcelIntegration.Metadata["PIPE_MASS"].IsThreadSafe, "reference and wizard guards must run on Excel thread");
        Check(ExcelIntegration.Arguments["PIPE_MASS"].Count == 32 && ((ExcelArgumentAttribute)ExcelIntegration.Arguments["PIPE_MASS"][0]).AllowReference, "expanded arguments carry real AllowReference metadata");
        Check(ExcelIntegration.Arguments["GKD_CUSTOM"].Count == 33, "universal packs name plus all 32 inputs");
        Check(Equals(Call(ExcelError.ExcelErrorDiv0, 1.0), ExcelError.ExcelErrorDiv0), "source errors propagate");
        Check(Call(1.0, ExcelMissing.Value).ToString().Contains("缺少"), "missing required parameter rejected");
        Check(Call(1.0, true).ToString().Contains("有限数值"), "booleans not silently coerced into numbers");
        Check(Call(1.0, double.NaN).ToString().Contains("有限数值"), "nonfinite numbers rejected");
        Check(Call(1.0).ToString().Contains("需要"), "arity mismatch rejected");
        var universal = new object[33]; for (int i = 0; i < 33; i++) universal[i] = ExcelMissing.Value; universal[0] = "pipe_mass"; universal[1] = 2.0; universal[2] = 5.0;
        Check(Equals(ExcelIntegration.Functions["GKD_CUSTOM"].DynamicInvoke(universal), 10.0), "universal Excel wrapper supports case-insensitive names");
        Check(Call(new object[,] { { 1.0, 2.0 } }, 3.0).ToString().Contains("单个数值"), "numeric deterministic inputs cannot consume a range implicitly");
        var reference = new ExcelReference { Values = new object[,] { { 3.0 } } };
        Check(Equals(Call(reference, reference), 9.0) && reference.Reads == 1, "same reference flattened once per invocation");
        var oversized = new ExcelReference { Oversized = true };
        Check(Call(oversized, 1.0).ToString().Contains("过大") && oversized.Reads == 0, "existing bounded reader rejects huge references before reads");
        ExcelDnaUtil.InWizard = true;
        Check(Equals(Call(reference, reference), ExcelError.ExcelErrorNA) && reference.Reads == 1 && GkdAiFunctions.Calls == 0, "wizard cannot read references or start AI");
        ExcelDnaUtil.InWizard = false;
        f.Enabled = false; GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { f });
        Check(Call(1.0, 2.0).ToString().Contains("停用"), "disabled current alias returns clear message");
        GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula>());
        Check(Call(1.0, 2.0).ToString().Contains("已删除"), "deleted current alias returns clear message");
        Check(GkdCustomFormulaFunctions.IsCustomName("pipe_mass") && GkdCustomFormulaFunctions.IsCustomName("gkd_custom") && !GkdCustomFormulaFunctions.IsCustomName("SUM"), "parent recognition includes registered tombstones");
        foreach (string name in new[] { "GKD", "GKD_CUSTOM", "SUM", "LET", "R", "XFD1048576" })
        { var collision = Program.Formula(name); Reject(() => GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { collision }), "reserved/cell collision rejected: " + name); }
        var ai = Program.Formula("PIPE_AI"); ai.Kind = GkdCustomFormulaKind.Ai; ai.FormulaExpression = ""; ai.PromptTemplate = "分析附表";
        GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { ai });
        Check(Equals(GkdCustomFormulaFunctions.Custom("PIPE_AI", new object[,] { { "a", "b" } }, "c"), "AI reply"), "AI executes through GKD runtime");
        Check(GkdAiFunctions.Matrix.GetLength(0) == 4 && Equals(GkdAiFunctions.Matrix[1, 0], "DIAMETER") && Equals(GkdAiFunctions.Matrix[2, 3], "b") && Equals(GkdAiFunctions.Matrix[3, 0], "LENGTH"), "AI matrix names and bounds explicit inputs");
        string first = GkdAiFunctions.Prompt;
        ai.Description = "Changed without version bump"; GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { ai });
        GkdCustomFormulaFunctions.Custom("PIPE_AI", new object[,] { { "a", "b" } }, "c"); Check(first != GkdAiFunctions.Prompt, "definition changes cannot reuse stale GKD cache");
        ai.Description = "质量计算"; GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { ai });
        GkdCustomFormulaFunctions.Custom("PIPE_AI", new object[,] { { "a", "b" } }, "c"); Check(first != GkdAiFunctions.Prompt, "A-B-A edit receives fresh request generation");
        var many = Program.Formula("MANY_INPUTS"); many.Inputs = Enumerable.Range(0, 32).Select(i => "ARG_" + i).ToList(); many.FormulaExpression = string.Join("+", many.Inputs);
        GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { many });
        var values = Enumerable.Repeat((object)1.0, 32).ToArray();
        Check(Equals(ExcelIntegration.Functions["MANY_INPUTS"].DynamicInvoke(values), 32.0), "all 32 actual registered slots evaluate");
        var tooMany = new ExcelReference { Values = new object[2500, 1], RowLast = 2499 };
        GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { ai });
        int aiCalls = GkdAiFunctions.Calls;
        Check(GkdCustomFormulaFunctions.Custom("PIPE_AI", tooMany, "c").ToString().Contains("10000") && tooMany.Reads == 0 && GkdAiFunctions.Calls == aiCalls, "aggregate AI size must be preflighted before reference reads");
        many.Inputs[0] = "NEW_ARGUMENT"; many.FormulaExpression = string.Join("+", many.Inputs);
        GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { many });
        Check(((ExcelArgumentAttribute)ExcelIntegration.Arguments["MANY_INPUTS"][0]).Name == "NEW_ARGUMENT", "saving argument names updates actual registration metadata");
        var r1c1 = Program.Formula("R1C1"); Reject(() => GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { r1c1 }), "R1C1 reference collision rejected");
        Check(ExcelIntegration.Metadata["MANY_INPUTS"].SuppressOverwriteError, "intentional metadata re-registration must not show Excel-DNA overwrite error dialogs");
        many.FormulaExpression = "1";
        Check(Equals(GkdCustomFormulaFunctions.Custom("MANY_INPUTS", values), 32.0), "caller mutations cannot silently mutate a validated runtime snapshot");
        var constant = Program.Formula("CONSTANT_VALUE"); constant.Inputs.Clear(); constant.FormulaExpression = "2^3";
        GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { constant });
        Check(Equals(GkdCustomFormulaFunctions.Custom("CONSTANT_VALUE"), 8.0), "zero input formulas evaluate safely");
        var divide = Program.Formula("DIVIDE_VALUES"); divide.FormulaExpression = "DIAMETER/LENGTH";
        GkdCustomFormulaFunctions.ApplyDefinitions(new List<GkdCustomFormula> { divide });
        Check(GkdCustomFormulaFunctions.Custom("DIVIDE_VALUES", 1.0, 0.0).ToString().StartsWith("GKD_CUSTOM：", StringComparison.Ordinal), "arithmetic errors become clear bounded formula errors");
        GkdCustomFormulaFunctions.Stop(); Check(GkdCustomFormulaFunctions.Custom("MANY_INPUTS", values).ToString().Contains("停止"), "shutdown prevents old delegates from executing");
        StoreFailures();
    }
    static void StoreFailures()
    {
        string root = Path.Combine(Path.GetTempPath(), "gkd-runtime-fail-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
        try
        {
            string path = Path.Combine(root, "formulas.json");
            GkdCustomFormulaStore.Save(new List<GkdCustomFormula> { Program.Formula() }, path);
            string before = File.ReadAllText(path);
            var collision = Program.Formula("SUM");
            Reject(() => GkdCustomFormulaStore.Save(new List<GkdCustomFormula> { collision }, path), "store must reject reserved names before persistence");
            Check(before == File.ReadAllText(path), "name collision leaves previous store intact");
            string blocked = Path.Combine(root, "blocked"); Directory.CreateDirectory(blocked);
            bool failed = false;
            try { GkdCustomFormulaStore.Save(new List<GkdCustomFormula> { Program.Formula() }, blocked); } catch (IOException) { failed = true; }
            Check(failed && Directory.Exists(blocked) && Directory.GetFiles(root, "*.tmp").Length == 0, "failed publication cleans temporary file and preserves existing target");
            var large = Program.Formula(); large.Description = new string('中', 4000);
            var oversized = new List<GkdCustomFormula>();
            for (int i = 0; i < 100; i++) { var item = Program.Formula("NAME_" + i); item.Description = large.Description; oversized.Add(item); }
            failed = false; try { GkdCustomFormulaStore.Save(oversized, path); } catch (InvalidDataException) { failed = true; }
            Check(failed && before == File.ReadAllText(path), "UTF8 byte cap enforced before touching old store");
            Parallel.For(0, 12, i => GkdCustomFormulaStore.Save(new List<GkdCustomFormula> { Program.Formula("NAME_" + i) }, path));
            Check(GkdCustomFormulaStore.Load(path).Count == 1 && Directory.GetFiles(root, "*.tmp").Length == 0, "concurrent saves publish complete valid snapshots");
        }
        finally { Directory.Delete(root, true); }
    }
}
