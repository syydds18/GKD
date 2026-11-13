using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Outline.Setup
{
    // Pure Windows-path parsing: registry commands are data, never executed.
    internal static class ExcelPathCandidates
    {
        private static bool IsAbsolute(string path)
        {
            return Regex.IsMatch(path, @"\A(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+\\)");
        }

        internal static string Executable(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            value = Environment.ExpandEnvironmentVariables(value.Trim());
            string path;
            if (value[0] == '"')
            {
                int end = value.IndexOf('"', 1);
                if (end < 0) return null;
                path = value.Substring(1, end - 1);
                string suffix = value.Substring(end + 1);
                if (suffix.Length > 0 && !char.IsWhiteSpace(suffix[0]) && suffix[0] != ',') return null;
            }
            else
            {
                var match = Regex.Match(value, @"\A(?<path>(?:[A-Za-z]:\\|\\\\)[^""\r\n]*?\\EXCEL\.EXE)(?=$|[\s,])", RegexOptions.IgnoreCase);
                if (!match.Success) return null;
                path = match.Groups["path"].Value;
            }
            return IsAbsolute(path) && path.EndsWith(@"\EXCEL.EXE", StringComparison.OrdinalIgnoreCase) ? path : null;
        }

        internal static IEnumerable<string> InstallRoot(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) yield break;
            value = Environment.ExpandEnvironmentVariables(value.Trim());
            if (value.Length >= 2 && value.StartsWith("\"", StringComparison.Ordinal) && value.EndsWith("\"", StringComparison.Ordinal))
                value = value.Substring(1, value.Length - 2);
            if (!IsAbsolute(value) || value.IndexOf('"') >= 0 || value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0) yield break;
            value = value.TrimEnd('\\');
            yield return value + @"\EXCEL.EXE";
            foreach (string version in new[] { "16", "15", "14" })
            {
                yield return value + @"\root\Office" + version + @"\EXCEL.EXE";
                yield return value + @"\Office" + version + @"\EXCEL.EXE";
            }
        }
    }
}
