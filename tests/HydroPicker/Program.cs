using System;
using System.IO;

static class Program
{
    static int Main()
    {
        string path = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "../../../../../src/Outline.Addin/HydropowerExtraction.cs"));
        string source = File.ReadAllText(path);
        string read = Between(source, "internal static List<HydroSheetChoice> ReadChoices", "private static List<string> ReadSelectedSheetNames");
        string form = Between(source, "internal sealed class HydropowerRangeForm", "internal static class HydropowerExtraction");
        Check(!read.Contains(".Load(workbookPath)"), "opening never reads the saved plan");
        Check(!read.Contains("savedRange"), "opening cannot override live print ranges");
        Check(read.Contains("HydroLivePrintArea.Read(app, sheet)"), "live range reader uses the supplied Excel instance");
        Check(form.Contains("使用提项目功能前，建议提前设置分页预览范围，并选择(可多选)需要导出的工作表范围。"), "exact picker hint");
        Check(form.Contains("Ui.Button(\"加载上次导表范围\")"), "explicit load button");
        Check(form.Contains("planActions.Controls.Add(savePlan); planActions.Controls.Add(loadPlan);"), "save/load buttons share header row");
        Check(form.Contains("HydroExportPlanPolicy.MatchVisibleRanges"), "load uses tested visible-sheet and range validation");
        int apply = form.IndexOf("private void LoadDefaultPlan()", StringComparison.Ordinal);
        Check(apply >= 0, "explicit load handler");
        string load = form.Substring(apply, form.IndexOf("private void SaveDefaultPlan()", apply, StringComparison.Ordinal) - apply);
        Check(load.IndexOf("matches.Count == 0", StringComparison.Ordinal) < load.IndexOf("row.Include =", StringComparison.Ordinal), "unusable plans return before changing current selections");
        Check(form.Contains("当前选择和范围未改变"), "unusable plan explains preserved selection");
        Check(form.Contains("private readonly List<HydroSheetChoice> workbookRows"), "picker retains workbook order separately from display order");
        Check(form.Contains("sheets.OrderByDescending(x => x.Include)"), "selected sheets appear first with stable ordering");
        Check(form.Contains("workbookRows.Where(x => x.Include && x.IsVisible).ToList()"), "output selection stays in workbook order and excludes hidden sheets");
        ReadChoicesTests.Run(source, path);
        Console.WriteLine("PASS picker source integration contract (live defaults, explicit safe load, hint/buttons, stable order)");
        return 0;
    }

    static string Between(string source, string first, string last)
    {
        int start = source.IndexOf(first, StringComparison.Ordinal);
        return source.Substring(start, source.IndexOf(last, start, StringComparison.Ordinal) - start);
    }
    static void Check(bool value, string message) { if (!value) throw new Exception("FAIL " + message); }
}
