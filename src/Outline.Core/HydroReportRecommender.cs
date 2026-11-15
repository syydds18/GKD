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
            var result = new HashSet<string>(StringComparer.Ordinal);
            foreach (string name in names ?? Enumerable.Empty<string>())
            {
                string n = Normalize(name);
                if (Excluded(n)) continue;
                if (IsCodeReport(n) || PhaseFor(n).Length != 0) result.Add(name ?? "");
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
            if (HasCodePrefix(codeName, "ZGS") || HasCodePrefix(codeName, "SN") || HasCodePrefix(codeName, "ZS")) return true;
            for (int i = 1; i <= 6; i++) if (HasCodePrefix(codeName, "G" + i)) return true;
            // Named report families only: a generic 概算/估算/匡算 title may be
            // a detail worksheet and is not sufficient on its own.
            return n.Contains("独立") || n.Contains("枢纽") || n.Contains("建设征地") || n.Contains("移民安置") ||
                n.Contains("补偿费用估算") || n.Contains("补偿费用概算") || n.Contains("补偿费用匡算");
        }

        private static bool Excluded(string n)
        {
            int offset = 0;
            while (offset < n.Length && char.IsDigit(n[offset])) offset++;
            return n.Substring(offset).StartsWith("FJ", StringComparison.Ordinal) ||
                n.Contains("资金流") || n.Contains("分年");
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
    }
}
