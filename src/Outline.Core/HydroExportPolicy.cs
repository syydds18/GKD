using System;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Outline.Core
{
    public static class HydroExportPolicy
    {
        public static string FileStem(DateTime day, string sheetName)
        {
            string safeName = new string((sheetName ?? "").Where(c => !char.IsControl(c) &&
                Array.IndexOf(Path.GetInvalidFileNameChars(), c) < 0 && "\\/:*?\"<>|".IndexOf(c) < 0).ToArray()).Trim();
            if (safeName.Length == 0) safeName = "水电工程";
            return day.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + " " + safeName + "-提项目";
        }

        public static bool IsTitleRow(string[] values, int rowOffset)
        {
            if (rowOffset < 0 || rowOffset >= 6 || values == null) return false;
            string title = null;
            foreach (string value in values)
            {
                if (string.IsNullOrWhiteSpace(value)) continue;
                if (title != null) return false;
                title = value.Trim();
            }
            if (title == null || title.Any(char.IsDigit)) return false;

            bool reportTitle = new[] { "总概算", "总估算", "总匡算", "概算表", "估算表", "匡算表", "资金流", "分年度" }
                .Any(term => title.Contains(term));
            bool projectTitle = title.Contains("工程") && new[] { "概算", "估算", "匡算", "费用", "投资" }
                .Any(term => title.Contains(term));
            return reportTitle || projectTitle;
        }
    }
}
