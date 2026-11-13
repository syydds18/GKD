using System;
using System.Collections.Generic;
using System.Linq;

namespace Outline.Core
{
    public static partial class Engine
    {
        private static readonly string[] HeaderKeys = { "编号", "序号", "项目编号", "编码" };
        private static readonly string[] NameKeys = { "名称", "项目名称", "名 称", "工程项目", "费用名称" };
        private static readonly string[] SectionTotalWords = { "一至", "静态投资", "基本预备费", "增值税", "总投资", "建设期利息" };
        public static Header FindHeader(SheetData sheet)
        { return FindHeader(sheet.Values, sheet.MaxRow, sheet.MaxCol); }
        private static Header FindHeader(Dictionary<int, Dictionary<int, object>> values, int maxRow, int maxCol)
        {
            Header best = null;
            for (int r = 1; r <= Math.Min(maxRow, 30); r++)
            {
                int codeCol = 0, nameCol = 0;
                for (int c = 1; c <= Math.Min(maxCol, 8); c++)
                {
                    string s = Norm(At(values, r, c)).Replace(" ", ""); if (s.Length == 0) continue;
                    if (codeCol == 0 && HeaderKeys.Any(s.Contains)) codeCol = c;
                    if (nameCol == 0 && NameKeys.Any(s.Contains)) nameCol = c;
                }
                if (codeCol > 0 && nameCol > 0) return new Header { HeaderRow = r, CodeCol = codeCol, NameCol = nameCol };
                if (codeCol > 0 && best == null) best = new Header { HeaderRow = r, CodeCol = codeCol, NameCol = codeCol + 1 };
            }
            if (best != null) return best;
            var counts = new SortedDictionary<int, int>();
            for (int c = 1; c <= Math.Min(maxCol, 4); c++)
            for (int r = 1; r <= Math.Min(maxRow, 200); r++)
            {
                string fam = CodeFamily(At(values, r, c)).Family;
                if (fam != null && fam != "part" && fam != "num") { if (!counts.ContainsKey(c)) counts[c] = 0; counts[c]++; }
            }
            int col = 1, max = 0;
            foreach (var pair in counts) if (pair.Value > max) { max = pair.Value; col = pair.Key; }
            return max < 3 ? new Header { HeaderRow = 3, CodeCol = 1, NameCol = 2, LowConfidence = true }
                : new Header { HeaderRow = 3, CodeCol = col, NameCol = col + 1 };
        }
        private static bool IsBannerRow(Dictionary<int, Dictionary<int, object>> values, int r, int nameCol, int maxCol)
        {
            for (int i = 0; i < 3; i++) { int c = nameCol + i; if (c >= 1 && c <= maxCol && CodeFamily(At(values, r, c)).Family == "part") return true; }
            return false;
        }
        public static Detection Detect(SheetData sheet)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            var values = sheet.Values ?? new Dictionary<int, Dictionary<int, object>>();
            int last = 0;
            for (int r = sheet.MaxRow; r >= 1; r--) if (!IsEmpty(values, r, Math.Max(1, sheet.MaxCol))) { last = r; break; }
            int maxRow = last != 0 ? last : sheet.MaxRow;
            Detection det = DetectLevels(values, maxRow, sheet.MaxCol);
            det.LastDataRow = maxRow;
            det.StructuredRows = det.Levels.Count(x => x.Value > 0);
            return det;
        }
        // The raw detectLevels variant intentionally does not trim trailing rows.
        public static Detection DetectLevels(SheetData sheet)
        { return DetectLevels(sheet.Values, sheet.MaxRow, sheet.MaxCol); }
        private sealed class OrdRecord { public int Row; public int Level; }
        private static Detection DetectLevels(Dictionary<int, Dictionary<int, object>> values, int maxRow, int maxCol)
        {
            Header header = FindHeader(values, maxRow, maxCol);
            var result = new Detection { HeaderRow = header.HeaderRow, CodeCol = header.CodeCol, NameCol = header.NameCol };
            if (header.LowConfidence) return result;
            int start = header.HeaderRow + 1, partCount = 0;
            var seen = new HashSet<string>();
            for (int r = start; r <= maxRow; r++)
            {
                string f = CodeFamily(At(values, r, header.CodeCol)).Family;
                if (f == "part") partCount++;
                if (f != null && seen.Add(f)) result.Families.Add(f);
            }
            bool partIsFamily = partCount >= 2;
            result.PartIsFamily = partIsFamily;
            foreach (string f in FamilyOrder) if (seen.Contains(f)) result.FamilyList.Add(f);
            var levels = result.Levels;
            int firstStruct = 0;
            for (int r = start; r <= maxRow; r++)
                if (IsBannerRow(values, r, header.NameCol, maxCol) || CodeFamily(At(values, r, header.CodeCol)).Family != null) { firstStruct = r; break; }
            if (firstStruct != 0) { for (int r = start; r < firstStruct; r++) levels[r] = 0; start = firstStruct; }
            var famLevel = new Dictionary<string, int>();
            var famRow = new Dictionary<string, int>();
            var famByOrd = new Dictionary<string, Dictionary<int, OrdRecord>>();
            int sectionStart = start, lastCodeLevel = 0, lastLevel = 0;
            for (int r = start; r <= maxRow; r++)
            {
                CodeInfo info = CodeFamily(At(values, r, header.CodeCol));
                string family = info.Family, nameVal = Norm(At(values, r, header.NameCol));
                if (IsBannerRow(values, r, header.NameCol, maxCol) || (family == "part" && !partIsFamily))
                {
                    levels[r] = 0; lastCodeLevel = 0; lastLevel = 0;
                    if (family == "part") sectionStart = r;
                    continue;
                }
                if (family != null)
                {
                    if (family == "roman" || family == "part" || family == "cn" || family == "cn_paren") sectionStart = r;
                    int lv;
                    if (family == "roman" || family == "part") lv = 0;
                    else if (family == "dotted")
                    {
                        int? first = info.Parts.Count > 0 ? (int?)info.Parts[0] : null;
                        int dotDepth = Math.Max(1, info.Parts.Count - 1);
                        string[] parentFams = { "num", "cn_paren", "cn", "roman", "part", "num_paren", "num_ju", "circle" };
                        int? baseLevel = null; int baseRow = -1;
                        if (first.HasValue)
                        foreach (string f in parentFams)
                        {
                            Dictionary<int, OrdRecord> records; OrdRecord rec;
                            if (famByOrd.TryGetValue(f, out records) && records.TryGetValue(first.Value, out rec)
                                && rec.Row < r && rec.Row >= sectionStart && rec.Row > baseRow)
                            { baseRow = rec.Row; baseLevel = rec.Level; }
                        }
                        if (!baseLevel.HasValue)
                        foreach (string f in parentFams)
                        {
                            int row;
                            if (famRow.TryGetValue(f, out row) && row < r && row >= sectionStart && row > baseRow)
                            { baseRow = row; baseLevel = famLevel[f]; }
                        }
                        if (!baseLevel.HasValue)
                        foreach (string f in parentFams)
                        {
                            int row;
                            if (famRow.TryGetValue(f, out row) && row < r && row > baseRow)
                            { baseRow = row; baseLevel = famLevel[f]; }
                        }
                        lv = Math.Min(MaxLevel, (baseLevel ?? 0) + dotDepth);
                    }
                    else
                    {
                        int recentRow = -1; int? parentLevel = null;
                        foreach (string f in FamilyParents[family])
                        {
                            int fl;
                            if (famLevel.TryGetValue(f, out fl) && famRow[f] >= recentRow)
                            { recentRow = famRow[f]; parentLevel = fl; }
                        }
                        lv = parentLevel.HasValue ? Math.Min(MaxLevel, parentLevel.Value + 1) : 0;
                    }
                    levels[r] = lv; famLevel[family] = lv; famRow[family] = r;
                    int? ord = CodeOrdinal(family, info.Text);
                    if (ord.HasValue)
                    {
                        if (!famByOrd.ContainsKey(family)) famByOrd[family] = new Dictionary<int, OrdRecord>();
                        famByOrd[family][ord.Value] = new OrdRecord { Row = r, Level = lv };
                    }
                    lastCodeLevel = lv; lastLevel = lv; continue;
                }
                if (IsEmpty(values, r, maxCol))
                {
                    bool laterHasData = false;
                    for (int rr = r + 1; rr <= Math.Min(maxRow, r + 30); rr++) if (!IsEmpty(values, rr, maxCol)) { laterHasData = true; break; }
                    levels[r] = laterHasData ? lastLevel : 0; continue;
                }
                int uncodedLevel;
                if (SectionTotalWords.Any(nameVal.Contains)) uncodedLevel = famLevel.ContainsKey("cn") ? famLevel["cn"] : 0;
                else uncodedLevel = Math.Min(MaxLevel, lastCodeLevel + 1);
                levels[r] = uncodedLevel; lastLevel = uncodedLevel;
            }
            result.MaxLevel = levels.Count == 0 ? 0 : levels.Values.Max();
            return result;
        }
    }
}
