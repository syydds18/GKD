using System;
using System.Linq;
using Outline.Core;

static class Program
{
    static int Main()
    {
        try
        {
            var plan = ProjectExtraction.Plan(ProjectType.Hydropower, 1, 1048576, 16384);
            var recommended = HydroReportRecommender.Recommend(new[] { "封面", "ZGS总概算", "SN独立", "G1", "G2", "G3", "G4", "G5", "G6", "分年度概算", "资金流概算表", "附表" });
            Check(recommended.SetEquals(new[] { "ZGS总概算", "SN独立", "G1", "G2", "G3", "G4", "G5", "G6" }), "recommend report families but exclude annual/cash-flow reports");
            Check(HydroReportRecommender.Recommend(new[] { "总估算", "甲表", "资金流估算表" }).SetEquals(new[] { "总估算" }), "do not select arbitrary sheets between estimate endpoints");
            Check(HydroReportRecommender.Recommend(new[] { "总匡算", "甲表", "分年度匡算表" }).SetEquals(new[] { "总匡算" }), "do not select arbitrary preliminary-budget details");
            Check(HydroReportRecommender.Recommend(new[] { "封面", "总概算", "资金流概算表", "备注" }).SetEquals(new[] { "总概算" }), "exclude cash-flow report after total report");
            Check(HydroReportRecommender.Recommend(new[] { "ZGS", "SN", "G1", "G2", "G3", "G4", "G5", "G6", "独立", "分年度", "资金流" }).SetEquals(new[] { "ZGS", "SN", "G1", "G2", "G3", "G4", "G5", "G6", "独立" }), "recommend standalone core hydro report names");
            Check(HydroReportRecommender.Recommend(new[] { "01-G1报告", "G10", "G100", "G11", "前言" }).SetEquals(new[] { "01-G1报告" }), "handle numeric sheet prefixes without matching G10-family names");
            Check(HydroReportRecommender.Recommend(new[] { "总估算", "无关说明", "资金流概算表", "封面", "总匡算", "甲表", "分年度匡算表" }).SetEquals(new[] { "总估算", "总匡算" }), "phase totals do not select adjacent detail sheets");
            Check(HydroReportRecommender.Recommend(new[] { "SN", "无关说明", "ZGS", "G10", "附录" }).SetEquals(new[] { "SN", "ZGS" }), "code reports remain recommended without phase endpoints");
            Check(HydroReportRecommender.Recommend(new[] { "前言", "总估算-项目", "甲表", "资金流估算表-汇总", "总概算（调整）", "乙表", "资金流概算表" }).SetEquals(new[] { "总估算-项目", "总概算（调整）" }), "recognize punctuated phase total reports conservatively");
            Check(HydroReportRecommender.Recommend(new[] { "枢纽工程概算表", "ZS", "02-ZS枢纽估算表", "建设征地补偿费用概算表", "移民安置补偿费用估算表", "补偿费用匡算表", "建设征地和移民安置" }).SetEquals(new[] { "枢纽工程概算表", "ZS", "02-ZS枢纽估算表", "建设征地补偿费用概算表", "移民安置补偿费用估算表", "补偿费用匡算表", "建设征地和移民安置" }), "recommend hub/ZS and G6 compensation aliases");
            Check(HydroReportRecommender.Recommend(new[] { "补偿费用估算", "补偿费用概算", "补偿费用匡算表" }).SetEquals(new[] { "补偿费用估算", "补偿费用概算", "补偿费用匡算表" }), "compensation phase aliases allow omitted table suffix");
            Check(HydroReportRecommender.Recommend(new[] { "FJ总概算", "02-FJ-G6建设征地", "fj独立费用", "ZS资金流概算", "G1分年表", "建设征地分年度概算表", "总概算", "建筑明细概算表", "设备安装概算表", "补偿费用明细表", "ZS10", "G60" }).SetEquals(new[] { "总概算" }), "exclusions win over aliases and generic detail tables stay unselected");
            Check(HydroExportPolicy.FileStem(new DateTime(2026, 11, 14), "报表/资金流:概算?*") == "20261114 报表资金流概算-提项目", "sanitize hydropower export file stem");
            Check(HydroExportPolicy.FileStem(new DateTime(2026, 1, 2), " /:*?\"<>| ") == "20260102 水电工程-提项目", "fallback when sanitized sheet name is empty");
            Check(HydroExportPolicy.IsTitleRow(new[] { "某某工程总概算" }, 0), "recognize first-row project report title");
            Check(HydroExportPolicy.IsTitleRow(new[] { "资金流概算表" }, 5), "recognize report title in first six rows");
            Check(!HydroExportPolicy.IsTitleRow(new[] { "某某工程总概算" }, 6), "do not classify titles beyond leading rows");
            Check(!HydroExportPolicy.IsTitleRow(new[] { "序号", "项目名称", "概算" }, 0), "do not classify multi-cell headers as titles");
            Check(!HydroExportPolicy.IsTitleRow(new[] { "1", "工程概算" }, 0), "do not classify numbered items as titles");
            Check(!HydroExportPolicy.IsTitleRow(new[] { "工程名称" }, 0), "do not classify generic engineering labels as report titles");
            Check(plan.DeleteColumnsFrom == 0, "hydropower does not apply fixed column pruning");
            Check(!plan.DeleteRowsAfterBlankInColumnC, "hydropower does not delete rows by blank numbering");
            Check(!plan.RatioDriven, "hydropower does not use ratio-driven pruning");
            Check(ProjectExtraction.Plan(ProjectType.Hydropower, 11, 648, 11).SheetIndex == 11, "hydropower supports arbitrary selected sheet order");
            Throws<ArgumentOutOfRangeException>(() => ProjectExtraction.Plan(ProjectType.Hydropower, 0, 120, 14));
            Throws<InvalidOperationException>(() => ProjectExtraction.Plan(ProjectType.Hydropower, 1, 0, 14));
            Throws<InvalidOperationException>(() => ProjectExtraction.Plan(ProjectType.Hydropower, 1, 120, 0));
            var heights = Enumerable.Repeat(15.0, 1024).ToArray();
            int reads = 0, writes = 0;
            HydroRowHeights.Pad(heights.Length, (first, last) =>
            {
                reads++;
                double h = heights[first - 1];
                return Enumerable.Range(first - 1, last - first + 1).All(i => heights[i] == h) ? (double?)h : null;
            }, (first, last, height) => { writes++; for (int i = first - 1; i < last; i++) heights[i] = height; });
            Check(reads == 1 && writes == 1 && heights.All(h => h == 22), "uniform 1024-row block padded with two operations");
            heights = new[] { 12.0, 40.0, 409.0, 15.0 };
            HydroRowHeights.Pad(heights.Length, (first, last) =>
                Enumerable.Range(first - 1, last - first + 1).All(i => heights[i] == heights[first - 1]) ? (double?)heights[first - 1] : null,
                (first, last, height) => { for (int i = first - 1; i < last; i++) heights[i] = height; });
            Check(heights.SequenceEqual(new[] { 20.0, 47.0, 409.0, 22.0 }), "mixed row heights retain individual padding and Excel maximum");
            Console.WriteLine("PASS hydropower-only project extraction plan");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("FAIL hydropower-only project extraction plan: " + ex.Message);
            return 1;
        }
    }

    static void Check(bool ok, string message) { if (!ok) throw new Exception(message); }
    static void Throws<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }
}
