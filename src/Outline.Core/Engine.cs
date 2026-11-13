using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Outline.Core
{
    public static partial class Engine
    {
        public const int MaxLevel = 7;
        private static readonly string[] FamilyOrder = { "part", "roman", "cn_paren", "cn", "num_paren", "num_ju", "dotted", "num", "circle" };
        // ECMAScript \d is ASCII; do not substitute .NET's Unicode \d.
        private static readonly string[] CodePatterns = {
            @"^第[一二三四五六七八九十百零〇0-9]{1,4}\s*(项|部分|章|节|篇)",
            @"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫⅬⅭⅮⅯ]{1,4}$", @"^[（(]\s*[一二三四五六七八九十百零〇]{1,3}\s*[)）]$",
            @"^[一二三四五六七八九十百零〇]{1,4}\s*[、.．,，]?$", @"^[（(]\s*[0-9]{1,3}\s*[)）]$",
            @"^[0-9]{1,3}\s*[)）]$", @"^[0-9]{1,2}(?:[.．][0-9]{1,2}){1,4}[.．]?$", @"^[0-9]{1,3}\s*[、.．]?$", @"^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]$"
        };
        private static readonly Dictionary<string, string[]> FamilyParents = new Dictionary<string, string[]> {
            {"roman", new string[0]}, {"part", new string[0]}, {"cn", new[]{"roman","part"}},
            {"cn_paren", new[]{"cn","roman","part"}}, {"num", new[]{"cn_paren","cn","roman","part"}},
            {"num_paren", new[]{"num","cn_paren","cn","roman","part"}},
            {"num_ju", new[]{"num_paren","num","cn_paren","cn","roman","part"}},
            {"dotted", new[]{"num","cn_paren","cn","roman","part"}},
            {"circle", new[]{"num_paren","num_ju","num","cn_paren","cn","roman","part"}}
        };
        // JS whitespace includes BOM but not U+0085. Keep norm/regex independent of OS culture.
        private const string JsSpace = "\u0009\u000a\u000b\u000c\u000d\u0020\u00a0\u1680\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u2028\u2029\u202f\u205f\u3000\ufeff";
        private static bool Test(string input, string pattern) { return Regex.IsMatch(input, pattern.Replace(@"\s", "[" + JsSpace + "]")); }
        private static string ReplaceSpace(string input, string replacement) { return Regex.Replace(input, "[" + JsSpace + "]+", replacement); }
        public static string Norm(object value)
        {
            if (value == null) return "";
            string s;
            if (value is bool) s = (bool)value ? "true" : "false";
            else if (IsNumeric(value)) s = JsNumberString(Convert.ToDouble(value, CultureInfo.InvariantCulture));
            else if (value is IFormattable) s = ((IFormattable)value).ToString(null, CultureInfo.InvariantCulture);
            else s = value.ToString();
            return s.Trim(JsSpace.ToCharArray()).Replace('\u3000', ' ').Replace('\u00a0', ' ');
        }
        private static bool IsNumeric(object value)
        { return value is double || value is float || value is decimal || value is int || value is long || value is short || value is byte || value is uint || value is ulong || value is ushort || value is sbyte; }
        private static string JsNumberString(double n)
        {
            if (n == 0) return "0";
            if (!Finite(n)) return double.IsNaN(n) ? "NaN" : n > 0 ? "Infinity" : "-Infinity";
            // Find the shortest round-tripping significand, including on .NET Framework,
            // where R can emit 17 digits even when a shorter JS-compatible form exists.
            double abs = Math.Abs(n); string text = null;
            for (int precision = 1; precision <= 17; precision++)
            {
                string candidate = abs.ToString("G" + precision, CultureInfo.InvariantCulture);
                double roundTrip;
                if (double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out roundTrip) && roundTrip == abs)
                { text = candidate; break; }
            }
            if (text == null) text = abs.ToString("R", CultureInfo.InvariantCulture);
            string[] split = text.ToUpperInvariant().Split('E');
            int exponent = split.Length > 1 ? int.Parse(split[1], CultureInfo.InvariantCulture) : 0;
            int dot = split[0].IndexOf('.');
            int position = (dot >= 0 ? dot : split[0].Length) + exponent;
            string digits = split[0].Replace(".", "");
            while (digits.Length > 1 && digits[0] == '0') { digits = digits.Substring(1); position--; }
            digits = digits.TrimEnd('0');
            string result;
            if (position > 0 && position <= 21)
                result = digits.Length <= position ? digits + new string('0', position - digits.Length) : digits.Insert(position, ".");
            else if (position <= 0 && position > -6) result = "0." + new string('0', -position) + digits;
            else
            {
                int e = position - 1;
                result = digits[0] + (digits.Length > 1 ? "." + digits.Substring(1) : "") + "e" + (e >= 0 ? "+" : "") + e.ToString(CultureInfo.InvariantCulture);
            }
            return n < 0 ? "-" + result : result;
        }
        private static bool Finite(double n) { return !double.IsNaN(n) && !double.IsInfinity(n); }
        public static string DecodeNumericCode(string s)
        {
            if (!Test(s, @"^[0-9]{1,2}\.[0-9]{3,}$")) return s;
            double n;
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out n) || !Finite(n)) return s;
            for (int k = 1; k <= 2; k++)
            {
                string t = n.ToString("F" + k, CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.');
                double parsed = double.Parse(t, CultureInfo.InvariantCulture);
                if (Math.Abs(parsed - n) < 1e-8 && Test(t, @"^[0-9]{1,2}(?:\.[0-9]{1,2})?$")) return t;
            }
            return s;
        }
        public static CodeInfo CodeFamily(object value)
        {
            string s = DecodeNumericCode(Norm(value));
            var result = new CodeInfo { Text = s };
            if (s.Length == 0 || s.Length > 40) return result;
            for (int i = 0; i < CodePatterns.Length; i++)
            {
                if (!Test(s, CodePatterns[i])) continue;
                result.Family = FamilyOrder[i]; result.Depth = 1;
                if (result.Family == "dotted")
                {
                    result.Depth = Math.Max(1, s.Count(c => c == '.' || c == '．'));
                    result.Parts = s.TrimEnd('.', '．').Split('.', '．').Select(x => int.Parse(x, CultureInfo.InvariantCulture)).ToList();
                }
                return result;
            }
            return result;
        }
        private static int? CnToNumber(string text)
        {
            string s = Regex.Replace(ReplaceSpace(Norm(text), ""), @"[、.．,，()（）]", "");
            if (s.Length == 0) return null;
            if (Test(s, @"^[0-9]+$")) return int.Parse(s, CultureInfo.InvariantCulture);
            const string digits = "零〇一二两三四五六七八九";
            int[] vals = { 0, 0, 1, 2, 2, 3, 4, 5, 6, 7, 8, 9 };
            int section = 0; int? number = null;
            foreach (char c in s)
            {
                int idx = digits.IndexOf(c);
                if (idx >= 0) number = vals[idx];
                else if (c == '十' || c == '百' || c == '千')
                { section += (number ?? 1) * (c == '十' ? 10 : c == '百' ? 100 : 1000); number = null; }
                else return null;
            }
            return section + (number ?? 0);
        }
        private static int? CodeOrdinal(string family, string text)
        {
            string s = Norm(text); if (s.Length == 0) return null;
            if (family == "part")
            {
                var m = Regex.Match(s, @"^第\s*([一二三四五六七八九十百零〇0-9]{1,4})\s*(?:项|部分|章|节|篇)");
                return m.Success ? CnToNumber(m.Groups[1].Value) : null;
            }
            if (family == "roman")
            { string roman = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ"; int i = roman.IndexOf(s[0]); return Test(s, "^[" + roman + "]{1,4}$") && i >= 0 ? (int?)(i + 1) : null; }
            if (family == "circle") { int i = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳".IndexOf(s[0]); return i >= 0 ? (int?)(i + 1) : null; }
            if (family == "cn" || family == "cn_paren") return CnToNumber(s);
            if (family == "num" || family == "num_paren" || family == "num_ju")
            { var m = Regex.Match(s, "[0-9]{1,3}"); return m.Success ? (int?)int.Parse(m.Value, CultureInfo.InvariantCulture) : null; }
            return null;
        }
        public static NumberValue ToNumberValue(object value)
        {
            if (value is double || value is float || value is decimal || value is int || value is long || value is short || value is byte || value is uint || value is ulong || value is ushort || value is sbyte)
            { double n = Convert.ToDouble(value, CultureInfo.InvariantCulture); return new NumberValue { Ok = Finite(n), Value = Finite(n) ? n : 0 }; }
            var result = new NumberValue();
            string s = value as string; if (s == null) return result;
            s = s.Trim(JsSpace.ToCharArray()); if (s.Length == 0) return result;
            if (Test(s, @"^#[A-Za-z0-9\/_!]+$")) { result.Error = s; return result; }
            bool neg = Test(s, @"^\(.*\)$"); if (neg) s = s.Substring(1, s.Length - 2);
            s = ReplaceSpace(s.Replace(",", "").Replace("，", ""), "");
            s = Regex.Replace(s, "^[￥¥]", "");
            if (!Test(s, @"^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$")) return result;
            double parsed;
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) || !Finite(parsed)) return result;
            result.Ok = true; result.Value = neg ? -parsed : parsed; return result;
        }
        public static string CellErrorValue(object value)
        {
            string s = value as string; if (s == null) return null;
            s = s.Trim(JsSpace.ToCharArray());
            if (!Test(s, @"^#[A-Za-z0-9\/_!]+$")) return null;
            return Regex.IsMatch(s, @"^#(REF|VALUE|DIV/0|NAME|N/A|NUM|NULL|SPILL|CALC|GETTING_DATA)!?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) ? s : null;
        }
        public static string ColLetter(int n)
        {
            string s = ""; while (n > 0) { int r = (n - 1) % 26; s = (char)(65 + r) + s; n = (n - 1) / 26; } return s;
        }
        private static object At(Dictionary<int, Dictionary<int, object>> values, int r, int c)
        { Dictionary<int, object> row; object v; return values != null && values.TryGetValue(r, out row) && row != null && row.TryGetValue(c, out v) ? v : null; }
        private static bool IsEmpty(Dictionary<int, Dictionary<int, object>> values, int r, int maxCol)
        { for (int c = 1; c <= maxCol; c++) { object v = At(values, r, c); if (v != null && !(v is string && (string)v == "")) return false; } return true; }
        private static int LevelAt(Dictionary<int, int> levels, int r) { int n; return levels != null && levels.TryGetValue(r, out n) ? n : 0; }
    }
}
