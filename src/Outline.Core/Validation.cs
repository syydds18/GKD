using System;
using System.Collections.Generic;
using System.Linq;

namespace Outline.Core
{
    public static partial class Engine
    {
        private static readonly string[] SumAmountKeys = { "合计", "金额", "费用", "投资", "造价", "含税", "税", "设备费", "安装费", "建筑工程费", "安装工程费", "设备购置费", "其他费用", "预备费", "静态投资", "总投资", "价差", "利息", "资金", "成本", "收入", "支出", "直接费", "间接费", "利润", "税金", "建筑及安装", "补偿" };
        private static readonly string[] SumExcludeKeys = { "单价", "数量", "单位", "费率", "比例", "税率", "定额", "指标", "系数", "备注", "序号", "编号", "名称", "日期", "年月", "%", "％" };
        private const string SubtotalPattern = @"小计|合计|总计|累计|总额|总概算|静态投资|总投资|建设投资|工程造价|合价|^一至";
        private const string FeePattern = @"增值税|基本预备费|价差预备费|建设期利息|预备费|^税金$";
        private const string SummaryPattern = @"分年度|完成工作量|年度完成|完成投资";
        private const string BannerPattern = @"^第[一二三四五六七八九十百零〇0-9]{1,4}\s*(项|部分|章|节|篇)";
        private const int MaxSumErrors = 500;
        public static bool IsAmountHeaderLabel(string label)
        {
            string s = ReplaceSpace(Norm(label), "");
            if (s.Length == 0 || SumExcludeKeys.Any(s.Contains)) return false;
            return SumAmountKeys.Any(s.Contains) || Test(s, @"第?\-?[0-9]+年");
        }
        private static bool IsSumBanner(object value)
        { string s = Norm(value); return s.Length > 0 && s.Length <= 40 && Test(s, BannerPattern); }
        private static List<int> SumHeaderRows(SheetData sheet, Detection detection, int last)
        {
            var rows = new List<int> { detection.HeaderRow };
            for (int r = detection.HeaderRow + 1; r <= Math.Min(detection.HeaderRow + 2, last); r++)
            {
                if (Norm(At(sheet.Values, r, detection.CodeCol)).Length > 0) break;
                bool banner = false;
                for (int i = 0; i < 3; i++) if (IsSumBanner(At(sheet.Values, r, detection.NameCol + i))) { banner = true; break; }
                if (banner) break;
                int textCount = 0;
                for (int c = 1; c <= sheet.MaxCol; c++)
                {
                    object v = At(sheet.Values, r, c);
                    if (v is string && Norm(v).Length > 0 && !ToNumberValue(v).Ok && CellErrorValue(v) == null) textCount++;
                }
                if (textCount < 2) break;
                rows.Add(r);
            }
            return rows;
        }
        public static ValidationResult Validate(SheetData sheet, Detection detection)
        {
            var empty = new ValidationResult(); if (sheet == null) return empty;
            detection = detection ?? new Detection();
            int last = detection.LastDataRow != 0 ? detection.LastDataRow : sheet.MaxRow;
            if (last == 0 || sheet.MaxCol == 0 || detection.HeaderRow == 0 || detection.CodeCol == 0 || detection.NameCol == 0)
            { empty.Reason = "未识别到表头或数据行"; return empty; }
            // Preserve JS's early exit: source outline alone never bypasses low-confidence detection.
            if (detection.MaxLevel <= 0 && detection.StructuredRows <= 0)
            { empty.Reason = "未识别到编号层级，无法判断父子关系。原表分级可能不是按编号生成的。"; return empty; }
            var headerRows = SumHeaderRows(sheet, detection, last);
            int headerEnd = headerRows[headerRows.Count - 1];
            var labels = new Dictionary<int, string>(); var amountCols = new List<int>();
            for (int c = 1; c <= sheet.MaxCol; c++)
            {
                labels[c] = string.Join(" ", headerRows.Select(r => Norm(At(sheet.Values, r, c))).Where(s => s.Length > 0));
                if (IsAmountHeaderLabel(labels[c])) amountCols.Add(c);
            }
            if (amountCols.Count == 0) { empty.Reason = "未找到金额列（数量、单价、费率等列已跳过）"; return empty; }
            if (sheet.OutlineMax > 0)
            {
                var result = Compute(sheet, detection, last, headerEnd, labels, amountCols, sheet.OutlineLevels, "outline");
                if (result.CheckedCells == 0 && detection.MaxLevel > 0 && detection.Levels != null)
                {
                    var detected = Compute(sheet, detection, last, headerEnd, labels, amountCols, detection.Levels, "detected");
                    if (detected.CheckedCells > 0) { detected.Fallback = true; return detected; }
                }
                return result;
            }
            return Compute(sheet, detection, last, headerEnd, labels, amountCols, detection.Levels, "detected");
        }
        private sealed class SumRow
        {
            public int Row, Level;
            public string Code, Name;
            public bool Banner, Sub, Fee, Summary;
            public Dictionary<int, double> Numbers = new Dictionary<int, double>();
            public Dictionary<int, string> Errors = new Dictionary<int, string>();
        }
        private sealed class Candidate
        { public List<int> List; public string Label; public double Sum; public int Count; }
        private static void AddCandidate(List<Candidate> candidates, HashSet<string> seen, IEnumerable<int> list, string label)
        {
            var uniq = list.Distinct().OrderBy(x => x).ToList();
            if (seen.Add(string.Join(",", uniq))) candidates.Add(new Candidate { List = uniq, Label = label });
        }
        private static void PushError(ValidationResult result, ValidationError error)
        { if (result.Errors.Count < MaxSumErrors) result.Errors.Add(error); else result.HiddenErrors++; }
        private static void SetFormula(SheetData sheet, ValidationError error)
        {
            Dictionary<int, string> row; string formula;
            error.HasFormula = sheet.Formulas != null && sheet.Formulas.TryGetValue(error.Row, out row) && row != null && row.ContainsKey(error.Col);
            error.Formula = error.HasFormula && sheet.Formulas[error.Row].TryGetValue(error.Col, out formula) ? formula ?? "" : "";
        }
        private static ValidationResult Compute(SheetData sheet, Detection detection, int last, int headerEnd, Dictionary<int, string> labels,
            List<int> amountCols, Dictionary<int, int> levels, string source)
        {
            var data = new List<SumRow>();
            for (int r = headerEnd + 1; r <= last; r++)
            {
                Dictionary<int, object> row;
                if (sheet.Values == null || !sheet.Values.TryGetValue(r, out row) || row == null) continue;
                string code = Norm(At(sheet.Values, r, detection.CodeCol)), name = Norm(At(sheet.Values, r, detection.NameCol));
                if (name.Length == 0) for (int i = 1; i <= 2; i++) { string v = Norm(At(sheet.Values, r, detection.NameCol + i)); if (v.Length > 0) { name = v; break; } }
                string label = name.Length > 0 ? name : code;
                var item = new SumRow { Row = r, Level = LevelAt(levels, r), Code = code, Name = label,
                    Banner = IsSumBanner(code) || IsSumBanner(name), Sub = Test(label, SubtotalPattern), Fee = Test(label, FeePattern), Summary = Test(label, SummaryPattern) };
                foreach (int c in amountCols)
                {
                    object raw = At(sheet.Values, r, c);
                    if (raw == null || (raw is string && (string)raw == "")) continue;
                    string err = CellErrorValue(raw);
                    if (err != null) { item.Errors[c] = err; continue; }
                    NumberValue num = ToNumberValue(raw); if (num.Ok) item.Numbers[c] = num.Value;
                }
                if (code.Length == 0 && name.Length == 0 && item.Numbers.Count == 0 && item.Errors.Count == 0) continue;
                data.Add(item);
            }
            // JS integer-key object enumeration is ascending, not insertion order.
            var children = new SortedDictionary<int, List<int>>(); var stack = new List<int>(); int bannerIdx = -1;
            for (int i = 0; i < data.Count; i++)
            {
                int lv = data[i].Level;
                if (data[i].Banner && lv == 0) { bannerIdx = i; stack.Clear(); stack.Add(i); continue; }
                if (lv == 0)
                {
                    if (bannerIdx >= 0) { if (!children.ContainsKey(bannerIdx)) children[bannerIdx] = new List<int>(); children[bannerIdx].Add(i); }
                    stack.Clear(); stack.Add(i); continue;
                }
                while (stack.Count > 0 && data[stack[stack.Count - 1]].Level >= lv) stack.RemoveAt(stack.Count - 1);
                int p = stack.Count > 0 ? stack[stack.Count - 1] : -1;
                if (p >= 0) { if (!children.ContainsKey(p)) children[p] = new List<int>(); children[p].Add(i); }
                stack.Add(i);
            }
            var result = new ValidationResult { LevelSource = source, MaxLevel = data.Count == 0 ? 0 : data.Max(x => x.Level) };
            foreach (var pair in children)
            {
                int p = pair.Key;
                if (pair.Value.Count == 0 || data[p].Sub) continue;
                var allKids = pair.Value.Where(k => !data[k].Sub).ToList(); if (allKids.Count == 0) continue;
                var coded = new List<int>(); var feeExtra = new List<int>(); var uncodedSummary = new HashSet<int>();
                foreach (int k in allKids)
                {
                    bool hasCode = CodeFamily(data[k].Code).Family != null;
                    if (hasCode) coded.Add(k);
                    if (data[k].Fee && !hasCode) feeExtra.Add(k);
                    if (data[k].Summary && !hasCode) uncodedSummary.Add(k);
                }
                var candidates = new List<Candidate>(); var seen = new HashSet<string>();
                if (coded.Count > 0)
                {
                    if (feeExtra.Count > 0) AddCandidate(candidates, seen, coded.Concat(feeExtra), "编号子项 + 费用行");
                    AddCandidate(candidates, seen, coded, "编号子项");
                    AddCandidate(candidates, seen, allKids.Where(k => !uncodedSummary.Contains(k)), "全部子项");
                }
                else AddCandidate(candidates, seen, allKids, "全部子项");
                bool rowChecked = false;
                foreach (int c in amountCols)
                {
                    double pv; if (!data[p].Numbers.TryGetValue(c, out pv)) continue;
                    var opts = new List<Candidate>();
                    foreach (var candidate in candidates)
                    {
                        double sum = 0; int count = 0;
                        foreach (int k in candidate.List) { double v; if (data[k].Numbers.TryGetValue(c, out v)) { sum += v; count++; } }
                        if (count >= 2) opts.Add(new Candidate { Sum = sum, Count = count, Label = candidate.Label, List = candidate.List });
                    }
                    if (opts.Count == 0) continue;
                    // Preserve the original heuristic: closest sum wins; ties retain the first candidate.
                    Candidate best = opts[0];
                    for (int oi = 1; oi < opts.Count; oi++) if (Math.Abs(pv - opts[oi].Sum) < Math.Abs(pv - best.Sum)) best = opts[oi];
                    result.CheckedCells++; rowChecked = true;
                    double diff = pv - best.Sum, tol = Math.Max(0.05, Math.Abs(pv) * 1e-9);
                    if (Math.Abs(diff) > tol)
                    {
                        if (Math.Abs(diff) > result.MaxDiff) result.MaxDiff = Math.Abs(diff);
                        string hint = ""; double matchTol = Math.Max(0.05, Math.Abs(diff) * 1e-9);
                        foreach (int k in best.List)
                        {
                            double kv; if (!data[k].Numbers.TryGetValue(c, out kv)) continue;
                            if (Math.Abs(Math.Abs(diff) - Math.Abs(kv)) <= matchTol)
                            { hint = "差额 = 子项「" + (data[k].Name.Length > 0 ? data[k].Name : "（无名称）") + "」，疑似重复计入或漏计"; break; }
                        }
                        var error = new ValidationError { Kind = "sum", Row = data[p].Row, Col = c, Letter = ColLetter(c),
                            RowName = data[p].Name.Length > 0 ? data[p].Name : "（无名称）", ColLabel = labels[c].Length > 0 ? labels[c] : ColLetter(c),
                            ParentValue = pv, ChildSum = best.Sum, Diff = diff, ChildCount = best.Count, Basis = best.Label, Hint = hint, Level = data[p].Level };
                        SetFormula(sheet, error); PushError(result, error);
                    }
                }
                if (rowChecked) result.CheckedRows++;
            }
            foreach (var item in data)
            foreach (int c in amountCols)
            {
                string err; if (!item.Errors.TryGetValue(c, out err)) continue;
                var error = new ValidationError { Kind = "cell", Row = item.Row, Col = c, Letter = ColLetter(c),
                    RowName = item.Name.Length > 0 ? item.Name : "（无名称）", ColLabel = labels[c].Length > 0 ? labels[c] : ColLetter(c),
                    ParentValue = null, ChildSum = null, Diff = 0, ChildCount = 0, Basis = "", Level = item.Level, Message = err };
                SetFormula(sheet, error); PushError(result, error);
            }
            result.Errors = result.Errors.OrderBy(e => e.Row).ThenBy(e => e.Col).ToList();
            result.ErrorGroups = result.Errors.GroupBy(e => e.Row).OrderBy(g => g.Key).Select(g => new ErrorGroup
                { Row = g.Key, RowName = g.First().RowName, Level = g.First().Level, Cells = g.ToList() }).ToList();
            result.ErrorRows = result.ErrorGroups.Count;
            result.ErrorCount = result.Errors.Count + result.HiddenErrors;
            result.Truncated = result.HiddenErrors > 0;
            result.Columns = amountCols.Select(c => new ColumnInfo { Col = c, Letter = ColLetter(c), Label = labels[c].Length > 0 ? labels[c] : ColLetter(c) }).ToList();
            result.Status = "pass"; result.Reason = "";
            if (result.Errors.Count > 0 || result.HiddenErrors > 0) result.Status = "diff";
            else if (result.CheckedCells == 0) { result.Status = "na"; result.Reason = "没有找到“父行 + 至少两个子项”的可校验结构"; }
            return result;
        }
    }
}
