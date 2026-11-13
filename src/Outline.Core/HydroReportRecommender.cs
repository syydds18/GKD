using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Outline.Core
{
    public static class HydroReportRecommender
    {
        private static readonly string[] PhaseStarts = { "总概算", "总估算", "总匡算" };

        public static HashSet<string> Recommend(IEnumerable<string> names)
        {
            var sheets = (names ?? Enumerable.Empty<string>())
                .Select((name, index) => new Entry(name ?? "", Normalize(name), index)).ToList();
            var result = new HashSet<string>(StringComparer.Ordinal);

            foreach (var sheet in sheets)
                if (IsCodeReport(sheet.Normalized) || PhaseFor(sheet.Normalized).Length != 0) result.Add(sheet.Name);

            for (int start = 0; start < sheets.Count; start++)
            {
                string phase = PhaseFor(sheets[start].Normalized);
                if (phase.Length == 0) continue;

                int nextStart = sheets.Count;
                for (int i = start + 1; i < sheets.Count; i++)
                    if (PhaseFor(sheets[i].Normalized).Length != 0) { nextStart = i; break; }

                int end = -1;
                for (int i = start + 1; i < nextStart; i++)
                    if (IsEnd(sheets[i].Normalized, phase)) { end = i; break; }
                if (end < 0) continue;
                for (int i = start; i <= end; i++) result.Add(sheets[i].Name);
            }
            return result;
        }

        public static string Normalize(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            var b = new StringBuilder(value.Length);
            foreach (char c in value)
                if (char.IsLetterOrDigit(c) || c >= '\u3400' && c <= '\u9fff') b.Append(char.ToUpperInvariant(c));
            return b.ToString();
        }

        private static bool IsCodeReport(string n)
        {
            int offset = 0;
            while (offset < n.Length && char.IsDigit(n[offset])) offset++;
            string codeName = n.Substring(offset);
            if (HasCodePrefix(codeName, "ZGS") || HasCodePrefix(codeName, "SN")) return true;
            for (int i = 1; i <= 6; i++) if (HasCodePrefix(codeName, "G" + i)) return true;
            return n.Contains("独立") || n.Contains("分年度") || n.Contains("资金流");
        }

        private static bool HasCodePrefix(string value, string code)
        {
            return value == code || value.StartsWith(code, StringComparison.Ordinal) &&
                (value.Length == code.Length || !char.IsDigit(value[code.Length]));
        }

        private static string PhaseFor(string n)
        {
            foreach (string phase in PhaseStarts)
                if (n.Contains(phase)) return phase.Substring(1);
            return "";
        }

        private static bool IsEnd(string n, string phase)
        {
            if (phase == "概算") return n.Contains("资金流概算");
            if (phase == "估算") return n.Contains("资金流估算");
            return n.Contains("分年度匡算");
        }

        private sealed class Entry
        {
            internal readonly string Name, Normalized;
            internal readonly int Index;
            internal Entry(string name, string normalized, int index) { Name = name; Normalized = normalized; Index = index; }
        }
    }
}
