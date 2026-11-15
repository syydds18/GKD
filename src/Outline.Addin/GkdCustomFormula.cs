using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Outline.Addin
{
    internal enum GkdCustomFormulaKind { Deterministic, Ai }

    internal static class GkdCustomFormulaJson
    {
        internal const int MaxJsonLength = 1048576;
        internal const int MaxFormulas = 256;
        private const string FormulaFields = "Id|ExcelName|DisplayName|Description|Kind|Inputs|FormulaExpression|PromptTemplate|OutputUnit|Version|Enabled";

        public static string Export(IEnumerable<GkdCustomFormula> formulas)
        {
            if (formulas == null) throw new ArgumentException("公式列表不能为空。");
            var array = new JArray();
            var ids = new HashSet<Guid>(); var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var f in formulas)
            {
                if (f == null) throw new ArgumentException("公式不能为空。");
                f.Validate();
                if (array.Count >= MaxFormulas || !ids.Add(Guid.Parse(f.Id)) || !names.Add(f.ExcelName)) throw new ArgumentException("公式数量过多，或标识、名称重复。");
                array.Add(new JObject
                {
                    ["Id"] = f.Id, ["ExcelName"] = f.ExcelName, ["DisplayName"] = f.DisplayName,
                    ["Description"] = f.Description, ["Kind"] = f.Kind.ToString(), ["Inputs"] = new JArray(f.Inputs),
                    ["FormulaExpression"] = f.FormulaExpression, ["PromptTemplate"] = f.PromptTemplate,
                    ["OutputUnit"] = f.OutputUnit, ["Version"] = f.Version, ["Enabled"] = f.Enabled
                });
            }
            string json = new JObject { ["SchemaVersion"] = 1, ["Formulas"] = array }.ToString(Formatting.Indented);
            if (json.Length > MaxJsonLength) throw new ArgumentException("公式文件过大。");
            return json;
        }

        public static List<GkdCustomFormula> Import(string json)
        {
            if (json == null || json.Length > MaxJsonLength) throw new InvalidDataException("公式文件为空或过大。");
            var root = GkdAiProtocol.ReadJson(json) as JObject;
            Fields(root, "SchemaVersion|Formulas");
            if (Integer(root["SchemaVersion"]) != 1 || !(root["Formulas"] is JArray)) throw new InvalidDataException("公式文件格式或版本无效。");
            if (((JArray)root["Formulas"]).Count > MaxFormulas) throw new InvalidDataException("公式数量过多。");
            var result = new List<GkdCustomFormula>();
            var ids = new HashSet<Guid>(); var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var token in (JArray)root["Formulas"])
            {
                var item = token as JObject; Fields(item, FormulaFields);
                foreach (string field in new[] { "Id", "ExcelName", "DisplayName", "Description", "Kind", "FormulaExpression", "PromptTemplate", "OutputUnit" })
                    if (item[field].Type != JTokenType.String) throw new InvalidDataException("公式文本字段类型无效。");
                if (item["Enabled"].Type != JTokenType.Boolean || !(item["Inputs"] is JArray)) throw new InvalidDataException("公式启用标记或输入列表类型无效。");
                foreach (var input in (JArray)item["Inputs"])
                    if (input.Type != JTokenType.String) throw new InvalidDataException("公式输入名称必须为文本。");
                string kind = (string)item["Kind"];
                if (kind != "Deterministic" && kind != "Ai") throw new InvalidDataException("公式类型无效。");
                var f = new GkdCustomFormula
                {
                    Id = (string)token["Id"], ExcelName = (string)token["ExcelName"], DisplayName = (string)token["DisplayName"],
                    Description = (string)token["Description"], Kind = kind == "Ai" ? GkdCustomFormulaKind.Ai : GkdCustomFormulaKind.Deterministic,
                    Inputs = token["Inputs"].ToObject<List<string>>(), FormulaExpression = (string)token["FormulaExpression"],
                    PromptTemplate = (string)token["PromptTemplate"], OutputUnit = (string)token["OutputUnit"], Version = Integer(token["Version"]), Enabled = (bool)token["Enabled"]
                };
                f.Validate();
                if (!ids.Add(Guid.Parse(f.Id)) || !names.Add(f.ExcelName)) throw new InvalidDataException("公式标识或名称重复。");
                result.Add(f);
            }
            return result;
        }
        private static void Fields(JObject item, string fields)
        {
            string[] required = fields.Split('|');
            if (item == null || item.Count != required.Length) throw new InvalidDataException("公式字段不完整或包含未知字段。");
            foreach (string field in required) if (item[field] == null) throw new InvalidDataException("公式字段不完整。");
        }
        private static int Integer(JToken token)
        {
            int value;
            if (token == null || token.Type != JTokenType.Integer || !int.TryParse(token.ToString(), out value)) throw new InvalidDataException("公式版本必须为整数。");
            return value;
        }
    }

    // Data only. These definitions are never compiled or evaluated as source code.
    internal sealed class GkdCustomFormula
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("D");
        public string ExcelName { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public string Description { get; set; } = "";
        public GkdCustomFormulaKind Kind { get; set; }
        public List<string> Inputs { get; set; } = new List<string>();
        public string FormulaExpression { get; set; } = "";
        public string PromptTemplate { get; set; } = "";
        public string OutputUnit { get; set; } = "";
        public int Version { get; set; } = 1;
        public bool Enabled { get; set; } = true;

        public void Validate()
        {
            Guid id;
            if (Id == null || Id.Length != 36 || !Guid.TryParseExact(Id, "D", out id) || id == Guid.Empty) throw new ArgumentException("公式标识必须为非空 GUID。");
            if (!Identifier(ExcelName)) throw new ArgumentException("公式名称必须符合 [A-Z][A-Z0-9_]{0,30}。");
            Text(DisplayName, 100, true); DescriptionText(Description, 4000); Text(OutputUnit, 32, false);
            if (Version < 1) throw new ArgumentException("公式版本必须为正整数。");
            if (Kind != GkdCustomFormulaKind.Deterministic && Kind != GkdCustomFormulaKind.Ai) throw new ArgumentException("公式类型无效。");
            if (Inputs == null || Inputs.Count > 32) throw new ArgumentException("公式最多支持32个输入。");
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (string name in Inputs)
                if (!Identifier(name) || !names.Add(name)) throw new ArgumentException("输入名称无效或重复。");
            if (Kind == GkdCustomFormulaKind.Deterministic)
            {
                Text(FormulaExpression, 4096, true);
                var symbols = new Dictionary<string, double>(StringComparer.Ordinal);
                foreach (string input in Inputs) symbols.Add(input, 0);
                new Expression(FormulaExpression, symbols, true).Parse();
                if (PromptTemplate != "") throw new ArgumentException("确定性公式不能包含 AI 提示词。");
            }
            else
            {
                if (FormulaExpression != "") throw new ArgumentException("AI 公式不能包含确定性表达式。");
                if (string.IsNullOrWhiteSpace(PromptTemplate) || PromptTemplate.Length > 16000) throw new ArgumentException("AI 提示词为空或过长。");
                foreach (char c in PromptTemplate)
                    if (char.IsControl(c) && c != '\r' && c != '\n' && c != '\t') throw new ArgumentException("AI 提示词包含无效字符。");
            }
        }
        public double Evaluate(IDictionary<string, double> values)
        {
            Validate();
            if (!Enabled || Kind != GkdCustomFormulaKind.Deterministic) throw new ArgumentException("公式未启用或不是确定性公式。");
            if (values == null || values.Count != Inputs.Count) throw new ArgumentException("必须显式提供全部输入。");
            var exact = new Dictionary<string, double>(StringComparer.Ordinal);
            foreach (var pair in values)
            {
                if (!Inputs.Contains(pair.Key) || double.IsNaN(pair.Value) || double.IsInfinity(pair.Value)) throw new ArgumentException("输入名称或数值无效。");
                exact.Add(pair.Key, pair.Value);
            }
            return new Expression(FormulaExpression, exact).Parse();
        }

        // Grammar: sum -> product -> unary -> power -> primary.
        // Exponentiation is right-associative and binds more tightly than unary minus.
        private sealed class Expression
        {
            private readonly string text;
            private readonly IDictionary<string, double> values;
            private readonly bool validateOnly;
            private int pos, depth;
            internal Expression(string text, IDictionary<string, double> values, bool validateOnly = false) { this.text = text; this.values = values; this.validateOnly = validateOnly; }
            internal double Parse()
            {
                double result = Sum(); Space(); if (pos != text.Length) Bad(); return result;
            }
            private void Space() { while (pos < text.Length && text[pos] == ' ') pos++; }
            private bool Take(char c) { Space(); if (pos < text.Length && text[pos] == c) { pos++; return true; } return false; }
            private void Need(char c) { if (!Take(c)) Bad(); }
            private double Sum()
            {
                double value = Product();
                while (true) { if (Take('+')) value = Finite(value + Product()); else if (Take('-')) value = Finite(value - Product()); else return value; }
            }
            private double Product()
            {
                double value = Unary();
                while (true)
                {
                    if (Take('*')) value = Finite(value * Unary());
                    else if (Take('/')) { double divisor = Unary(); if (!validateOnly && divisor == 0) Bad(); value = Finite(validateOnly ? 0 : value / divisor); }
                    else return value;
                }
            }
            private double Unary()
            {
                if (++depth > 64) Bad();
                try { if (Take('+')) return Unary(); if (Take('-')) return -Unary(); return Power(); }
                finally { depth--; }
            }
            private double Power() { double value = Primary(); return Take('^') ? Finite(Math.Pow(value, Unary())) : value; }
            private double Primary()
            {
                if (Take('(')) { double value = Sum(); Need(')'); return value; }
                Space(); int start = pos;
                if (pos < text.Length && text[pos] >= 'A' && text[pos] <= 'Z')
                {
                    while (pos < text.Length && ((text[pos] >= 'A' && text[pos] <= 'Z') || (text[pos] >= '0' && text[pos] <= '9') || text[pos] == '_')) pos++;
                    string name = text.Substring(start, pos - start); double value;
                    if (name == "CONVERT" && Take('('))
                    {
                        double source = Sum(); Need(','); string from = UnitName(); Need(','); string to = UnitName(); Need(')');
                        Unit first, second;
                        if (!Units.TryGetValue(from, out first) || !Units.TryGetValue(to, out second) || first.Dimension != second.Dimension) Bad();
                        // Ratio first avoids a spurious overflow when converting within a large-scale unit.
                        return Finite(source * (Units[from].Scale / Units[to].Scale));
                    }
                    if (!values.TryGetValue(name, out value)) Bad(); return value;
                }
                bool digits = Digits();
                if (pos < text.Length && text[pos] == '.') { pos++; digits = Digits() || digits; }
                if (!digits) Bad();
                if (pos < text.Length && (text[pos] == 'e' || text[pos] == 'E'))
                {
                    pos++; if (pos < text.Length && (text[pos] == '+' || text[pos] == '-')) pos++;
                    if (!Digits()) Bad();
                }
                double number;
                if (!double.TryParse(text.Substring(start, pos - start), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number) || double.IsNaN(number) || double.IsInfinity(number)) Bad();
                return Finite(number);
            }
            private bool Digits() { int start = pos; while (pos < text.Length && text[pos] >= '0' && text[pos] <= '9') pos++; return start != pos; }
            private string UnitName()
            {
                Need('"'); int start = pos;
                while (pos < text.Length && text[pos] != '"')
                {
                    char c = text[pos++];
                    if (!((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')) || pos - start > 8) Bad();
                }
                string unit = text.Substring(start, pos - start); Need('"'); return unit;
            }
            private sealed class Unit
            {
                internal readonly string Dimension; internal readonly double Scale;
                internal Unit(string dimension, double scale) { Dimension = dimension; Scale = scale; }
            }
            private static readonly Dictionary<string, Unit> Units = new Dictionary<string, Unit>(StringComparer.Ordinal)
            {
                {"mm",new Unit("length",0.001)}, {"cm",new Unit("length",0.01)}, {"m",new Unit("length",1)}, {"km",new Unit("length",1000)},
                {"mm2",new Unit("area",0.000001)}, {"cm2",new Unit("area",0.0001)}, {"m2",new Unit("area",1)}, {"km2",new Unit("area",1000000)}, {"ha",new Unit("area",10000)},
                {"mm3",new Unit("volume",0.000000001)}, {"cm3",new Unit("volume",0.000001)}, {"m3",new Unit("volume",1)}, {"L",new Unit("volume",0.001)},
                {"g",new Unit("mass",0.001)}, {"kg",new Unit("mass",1)}, {"t",new Unit("mass",1000)},
                {"s",new Unit("time",1)}, {"min",new Unit("time",60)}, {"h",new Unit("time",3600)},
                {"Pa",new Unit("pressure",1)}, {"kPa",new Unit("pressure",1000)}, {"MPa",new Unit("pressure",1000000)},
                {"N",new Unit("force",1)}, {"kN",new Unit("force",1000)}
            };
            private double Finite(double value) { if (validateOnly) return 0; if (double.IsNaN(value) || double.IsInfinity(value)) Bad(); return value; }
            private static void Bad() { throw new ArgumentException("表达式包含无效语法、输入或数值运算。"); }
        }

        internal static bool Identifier(string value)
        {
            return value != null && Regex.IsMatch(value, @"\A[A-Z][A-Z0-9_]{0,30}\z", RegexOptions.CultureInvariant);
        }
        internal static void Text(string value, int max, bool required)
        {
            if (value == null || value.Length > max || (required && string.IsNullOrWhiteSpace(value))) throw new ArgumentException("公式文本为空或过长。");
            foreach (char c in value) if (char.IsControl(c)) throw new ArgumentException("公式文本包含控制字符。");
        }
        internal static void DescriptionText(string value, int max)
        {
            if (value == null || value.Length > max || string.IsNullOrWhiteSpace(value)) throw new ArgumentException("公式含义说明为空或过长。");
            foreach (char c in value)
                if (char.IsControl(c) && c != '\r' && c != '\n' && c != '\t') throw new ArgumentException("公式含义说明包含无效控制字符。");
        }
    }
}
