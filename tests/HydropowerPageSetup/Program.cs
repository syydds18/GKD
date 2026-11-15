using System;
using System.IO;

static class Program
{
    static int Main()
    {
        string source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../src/Outline.Addin/HydropowerExtraction.cs");
        source = Path.GetFullPath(source);
        string text = File.ReadAllText(source);
        Check(text.Contains("target.PageSetup.Zoom = false;"), "fit-to-pages mode disables fixed zoom");
        Check(text.Contains("target.PageSetup.FitToPagesWide = 1;"), "fit width to one page");
        Check(text.Contains("target.PageSetup.FitToPagesTall = false;"), "fit height is unlimited via Boolean false");
        Check(!text.Contains("target.PageSetup.FitToPagesTall = 0;"), "zero is not passed as page count");
        Check(text.Contains("HydroRangeCopy.CopyValuesAndFormats((object)sourceRange, (object)dest);"), "values are committed through tested native Excel transfer");
        Check(!text.Contains("dest.Value2 = dest.Value2;"), "COM error Variants are not reassigned as numeric values");
        Check(!text.Contains("for (int r = 1; r <= (int)sourceRange.Rows.Count; r++)"), "source row heights are not copied one COM call at a time");
        Check(text.Contains("range.Rows.AutoFit();"), "row heights are auto-fitted in a single Excel operation");
        if (!text.Contains("HydroExportAppearance.ClearFills((object)target);")) throw new Exception("all exported cells including conditional fill must be cleared");
        string runner = File.ReadAllText(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../../src/Outline.Addin/ProjectExtractionRunner.cs")));
        Check(runner.Contains("Path.GetFileNameWithoutExtension(sourcePath)"), "export name uses the workbook filename");
        Check(runner.Contains("exportWorkbookName, progress"), "export uses captured workbook filename");
        Console.WriteLine("PASS hydropower PageSetup fit settings");
        return 0;
    }
    static void Check(bool ok, string name) { if (!ok) throw new Exception("FAIL " + name); }
}
