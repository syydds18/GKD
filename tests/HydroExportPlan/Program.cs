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
            Require(!HydroExportPlanPolicy.ShouldInclude("Appendix", false, null, loaded), "opening never silently selects a saved sheet");
            Require(HydroExportPlanPolicy.ShouldInclude("Other", true, null, loaded), "saved plan does not override opening recommendations");
            Require(HydroExportPlanPolicy.ShouldInclude("Summary", true, null, null), "recommendations are fallback without plan/selection");
            string planPath = Directory.GetFiles(Path.Combine(root, "plans")).Single();
            File.WriteAllText(planPath, "<Other><Sheet name='Summary' range='A1:D20'/></Other>");
            Require(store.Load(workbook) == null, "unrelated XML is not an export plan");
            File.WriteAllText(planPath, "<HydroExportPlan version='99'><Sheet name='Summary' range='A1:D20'/></HydroExportPlan>");
            Require(store.Load(workbook) == null, "unsupported plan version is rejected");
            File.WriteAllText(planPath, "<HydroExportPlan");
            Require(store.Load(workbook) == null, "malformed XML is ignored safely");

            loaded.Sheets.Add(new HydroExportPlanSheet { Name = "RenamedOld", PrintRange = "A1:B2" });
            loaded.Sheets.Add(new HydroExportPlanSheet { Name = "Hidden", PrintRange = "A1:B2" });
            loaded.Sheets.Add(new HydroExportPlanSheet { Name = "Bad", PrintRange = "B2:A1" });
            loaded.Sheets.Add(new HydroExportPlanSheet { Name = "Huge", PrintRange = new string('Z', 80) + "1" });
            loaded.Sheets.Add(new HydroExportPlanSheet { Name = "WrongSheet", PrintRange = "'Summary'!A1:B2" });
            loaded.Sheets.Add(new HydroExportPlanSheet { Name = "BadQuote", PrintRange = "'!A1" });
            loaded.Sheets.Add(new HydroExportPlanSheet { Name = "O'Brien!", PrintRange = "='O''Brien!'!$B$2:$D$10" });
            loaded.Sheets.Add(new HydroExportPlanSheet { Name = "Summary", PrintRange = "A1" });
            var matched = HydroExportPlanPolicy.MatchVisibleRanges(loaded,
                new[] { "Appendix", "Summary", "RenamedNew", "Bad", "Huge", "WrongSheet", "BadQuote", "O'Brien!" });
            Require(matched.Select(x => x.Name).SequenceEqual(new[] { "Appendix", "Summary", "O'Brien!" }), "load only valid visible exact-name matches in workbook order");
            Require(matched[1].PrintRange == "A1:D20", "duplicate entries keep first valid saved range");
            Require(HydroExportPlanPolicy.MatchVisibleRanges(loaded, new[] { "RenamedNew" }).Count == 0, "renamed plans have no accidental matches");
            Require(HydroExportPlanPolicy.MatchVisibleRanges(null, new[] { "Summary" }).Count == 0, "missing plan has no matches");
            Console.WriteLine("PASS workbook-scoped XML persistence, opening precedence, explicit validated plan matching");
            return 0;
        }
        finally { Directory.Delete(root, true); }
    }

    private static void Require(bool value, string message)
    {
        if (!value) throw new Exception("FAIL " + message);
    }
}