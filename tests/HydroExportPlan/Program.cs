using System;
using System.IO;
using System.Linq;
using Outline.Core;

static class Program
{
    static int Main()
    {
        string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "hydro-plan-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var store = new HydroExportPlanStore(Path.Combine(root, "plans"));
            string workbook = Path.Combine(root, "book.xlsx");
            store.Save(workbook, new[] {
                new HydroExportPlanSheet { Name = "Summary", PrintRange = "A1:D20" },
                new HydroExportPlanSheet { Name = "Appendix", PrintRange = "B2:F30" }
            });

            var loaded = store.Load(workbook);
            Require(loaded != null && loaded.Sheets.Count == 2, "round trip sheet count");
            Require(loaded.Sheets[0].PrintRange == "A1:D20", "round trip range");
            Require(store.Load(Path.Combine(root, "other.xlsx")) == null, "workbook path isolation");

            Require(HydroExportPlanPolicy.ShouldInclude("Appendix", true, new[] { "Appendix", "Data" }, loaded), "Excel multi-selection includes selected sheet");
            Require(!HydroExportPlanPolicy.ShouldInclude("Summary", true, new[] { "Appendix", "Data" }, loaded), "Excel multi-selection overrides saved/default recommendations");
            Require(HydroExportPlanPolicy.ShouldInclude("Appendix", false, null, loaded), "saved plan overrides recommendations");
            Require(!HydroExportPlanPolicy.ShouldInclude("Other", true, null, loaded), "saved plan excludes non-saved sheet");
            Require(HydroExportPlanPolicy.ShouldInclude("Summary", true, null, null), "recommendations are fallback without plan/selection");
            Console.WriteLine("PASS workbook-scoped XML persistence and selection precedence");
            return 0;
        }
        finally { Directory.Delete(root, true); }
    }

    private static void Require(bool value, string message)
    {
        if (!value) throw new Exception("FAIL " + message);
    }
}