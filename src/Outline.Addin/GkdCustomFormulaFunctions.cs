using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using ExcelDna.Integration;

namespace Outline.Addin
{
    // All Excel operations run on the caller's Excel thread. Only declarative numeric
    // expressions reach the evaluator; the compiled lambdas below only pack UDF arguments.
    public static class GkdCustomFormulaFunctions
    {
        private static readonly object Gate = new object();
        private static Dictionary<string, GkdCustomFormula> current = new Dictionary<string, GkdCustomFormula>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> registered = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> registrationMetadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> reserved = BuildReservedNames();
        private static bool universalRegistered, running;
        private static long generation;
        private static string snapshot = "";
        private static string failure;
        private static readonly string session = Guid.NewGuid().ToString("N");
        private const string Prefix = "GKD_CUSTOM：";
        private delegate object AliasCall(object a01, object a02, object a03, object a04, object a05, object a06, object a07, object a08, object a09, object a10, object a11, object a12, object a13, object a14, object a15, object a16, object a17, object a18, object a19, object a20, object a21, object a22, object a23, object a24, object a25, object a26, object a27, object a28, object a29, object a30, object a31, object a32);
        private delegate object UniversalCall(object name, object a01, object a02, object a03, object a04, object a05, object a06, object a07, object a08, object a09, object a10, object a11, object a12, object a13, object a14, object a15, object a16, object a17, object a18, object a19, object a20, object a21, object a22, object a23, object a24, object a25, object a26, object a27, object a28, object a29, object a30, object a31, object a32);

        // Call from AutoOpen / a main-thread UI save callback, never Task.Run.
        internal static void Start() { Reload(); }
        internal static void Reload()
        {
            try { ApplyDefinitions(GkdCustomFormulaStore.Load()); }
            catch
            {
                lock (Gate) { current = new Dictionary<string, GkdCustomFormula>(StringComparer.OrdinalIgnoreCase); failure = "公式库读取或注册失败，请修复后重新加载。"; generation++; running = true; }
                EnsureUniversal();
                throw;
            }
        }
        internal static void Stop()
        {
            lock (Gate) { running = false; current = new Dictionary<string, GkdCustomFormula>(StringComparer.OrdinalIgnoreCase); generation++; snapshot = ""; failure = null; }
            // Excel-DNA owns delegate lifetimes until the XLL unloads. Keep registered
            // names as tombstones so a previously registered alias never calls a stale closure.
        }
        internal static bool IsAiFormula(string name)
        {
            if (name == null) return false;
            lock (Gate)
            {
                GkdCustomFormula formula;
                return current.TryGetValue(name, out formula) && formula.Enabled && formula.Kind == GkdCustomFormulaKind.Ai;
            }
        }
        internal static bool IsCustomName(string name)
        {
            if (name == null) return false;
            lock (Gate) return string.Equals(name, "GKD_CUSTOM", StringComparison.OrdinalIgnoreCase) || current.ContainsKey(name) || registered.Contains(name);
        }
        internal static void ValidateDefinitions(List<GkdCustomFormula> definitions)
        {
            GkdCustomFormulaJson.Export(definitions);
            foreach (var f in definitions)
            {
                if (reserved.Contains(f.ExcelName) || IsCellName(f.ExcelName))
                    throw new ArgumentException("公式名称与 Excel 或加载项保留名称冲突：" + f.ExcelName);
            }
        }
        internal static void ApplyDefinitions(List<GkdCustomFormula> definitions)
        {
            ValidateDefinitions(definitions);
            string json = GkdCustomFormulaJson.Export(definitions);
            var copies = GkdCustomFormulaJson.Import(json).ToDictionary(f => f.ExcelName, StringComparer.OrdinalIgnoreCase);
            lock (Gate)
            {
                if (!string.Equals(snapshot, json, StringComparison.Ordinal)) generation++;
                current = copies; snapshot = json; failure = null; running = true;
            }
            EnsureUniversal();
            foreach (var formula in copies.Values)
            {
                if (!formula.Enabled) continue;
                string metadata = formula.DisplayName + "\n" + formula.Description + "\n" + string.Join(",", formula.Inputs);
                lock (Gate) { string previous; if (registrationMetadata.TryGetValue(formula.ExcelName, out previous) && previous == metadata) continue; }
                Register(formula.ExcelName, formula);
                lock (Gate) { registered.Add(formula.ExcelName); registrationMetadata[formula.ExcelName] = metadata; }
            }
        }
        private static void EnsureUniversal()
        {
            lock (Gate) { if (universalRegistered) return; }
            Register("GKD_CUSTOM", null);
            lock (Gate) universalRegistered = true;
        }
        private static void Register(string name, GkdCustomFormula formula)
        {
            bool universal = formula == null;
            int count = universal ? 33 : 32;
            var parameters = Enumerable.Range(0, count).Select(i => Expression.Parameter(typeof(object), "p" + i)).ToArray();
            // RegisterDelegates does NOT expand params object[]. Explicitly produce
            // a fixed scalar signature compatible with its v1.9 loader.
            var packed = Expression.NewArrayInit(typeof(object), parameters.Skip(universal ? 1 : 0));
            var body = Expression.Call(typeof(GkdCustomFormulaFunctions).GetMethod("Custom"), universal ? (Expression)parameters[0] : Expression.Constant(name, typeof(object)), packed);
            Delegate wrapper = Expression.Lambda(universal ? typeof(UniversalCall) : typeof(AliasCall), body, parameters).Compile();
            var arguments = new List<object>();
            if (universal) arguments.Add(new ExcelArgumentAttribute { Name = "名称", Description = "已保存并启用的自定义公式 Excel 标识", AllowReference = false });
            for (int i = 0; i < 32; i++)
                arguments.Add(new ExcelArgumentAttribute { Name = formula != null && i < formula.Inputs.Count ? formula.Inputs[i] : "参数" + (i + 1), Description = "按定义顺序输入；最多32个参数；仅读取显式引用", AllowReference = true });
            ExcelIntegration.RegisterDelegates(new List<Delegate> { wrapper }, new List<object> { new ExcelFunctionAttribute { Name = name, Category = "搞快点 自定义公式", Description = universal ? "调用本机保存的自定义公式：GKD_CUSTOM(名称,参数...)" : formula.DisplayName + "：" + formula.Description, IsMacroType = true, IsThreadSafe = false, SuppressOverwriteError = true } }, new List<List<object>> { arguments });
        }
        // The .dna permits implicit exports: suppress this raw paramarray entry.
        // Only the scalar wrappers above are sent to RegisterDelegates.
        [ExcelFunction(Name = "GKD_CUSTOM", ExplicitRegistration = true)]
        public static object Custom(object nameValue, params object[] args)
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNA;
            try
            {
                if (nameValue is ExcelError) return nameValue;
                string name = nameValue as string;
                if (string.IsNullOrEmpty(name)) return Prefix + "请输入公式名称。";
                GkdCustomFormula definition; long revision;
                lock (Gate)
                {
                    if (!running) return Prefix + "运行时已停止。";
                    if (failure != null) return Prefix + failure;
                    if (!current.TryGetValue(name, out definition)) return Prefix + "公式不存在或已删除：" + name;
                    if (!definition.Enabled) return Prefix + "公式已停用：" + name;
                    revision = generation;
                }
                if (args == null) args = new object[0];
                if (args.Length > 32) return Prefix + "最多支持32个参数。";
                int supplied = args.Length;
                while (supplied > definition.Inputs.Count && args[supplied - 1] is ExcelMissing) supplied--;
                if (supplied != definition.Inputs.Count) return Prefix + "需要 " + definition.Inputs.Count + " 个参数。";
                // Preflight geometry before any cell reads, including aggregate limits.
                long planned = 0;
                for (int i = 0; i < supplied; i++)
                {
                    var reference = args[i] as ExcelReference;
                    var array = args[i] as object[,];
                    planned += reference != null ? ReferenceCells(reference) : array != null ? array.LongLength : 1;
                    if (planned > 10000 || (definition.Kind == GkdCustomFormulaKind.Ai && (planned + 1) * 4 > 10000))
                        return Prefix + "参数展开后超过10000个单元格，请缩小引用范围。";
                }
                var values = new object[supplied];
                int cells = 0;
                var references = new Dictionary<ExcelReference, object[,]>();
                for (int i = 0; i < supplied; i++)
                {
                    object value = args[i];
                    if (value is ExcelError) return value;
                    if (value == null || value is ExcelMissing || value is ExcelEmpty) return Prefix + "缺少参数：" + definition.Inputs[i];
                    var reference = value as ExcelReference;
                    if (reference != null)
                    {
                        object[,] read;
                        if (!references.TryGetValue(reference, out read)) { read = GkdAiFunctions.ReadReference(reference); references.Add(reference, read); }
                        value = read;
                    }
                    var array = value as object[,];
                    cells += array == null ? 1 : array.Length;
                    if (cells > 10000 || (array != null && array.Length == 0)) return Prefix + "参数合计应为1至10000个单元格。";
                    if (array != null)
                        foreach (object item in array) if (item is ExcelError) return item;
                    values[i] = value;
                }
                if (definition.Kind == GkdCustomFormulaKind.Deterministic)
                {
                    var numbers = new Dictionary<string, double>(StringComparer.Ordinal);
                    for (int i = 0; i < supplied; i++)
                    {
                        object value = values[i]; var array = value as object[,];
                        if (array != null) { if (array.Length != 1) return Prefix + "确定性公式的参数必须是单个数值：" + definition.Inputs[i]; value = array[array.GetLowerBound(0), array.GetLowerBound(1)]; }
                        double number;
                        if (value is bool || value == null || value is ExcelEmpty || value is ExcelMissing || !double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out number) || double.IsNaN(number) || double.IsInfinity(number))
                            return Prefix + "参数必须是有限数值：" + definition.Inputs[i];
                        numbers.Add(definition.Inputs[i], number);
                    }
                    return definition.Evaluate(numbers);
                }
                // Long form matrix preserves each explicit argument's identity and
                // row/column position without introducing unbounded rectangle padding.
                var matrix = new object[cells + 1, 4];
                matrix[0, 0] = "参数"; matrix[0, 1] = "行"; matrix[0, 2] = "列"; matrix[0, 3] = "值";
                // Header also counts toward GKD's 10,000-cell input limit.
                if ((long)(cells + 1) * 4 > 10000) return Prefix + "AI 参数展开后超过10000个单元格，请缩小引用范围。";
                int row = 1;
                for (int i = 0; i < supplied; i++)
                {
                    var array = values[i] as object[,];
                    int rows = array == null ? 1 : array.GetLength(0), cols = array == null ? 1 : array.GetLength(1);
                    for (int r = 0; r < rows; r++) for (int c = 0; c < cols; c++)
                    {
                        matrix[row, 0] = definition.Inputs[i]; matrix[row, 1] = r + 1; matrix[row, 2] = c + 1;
                        object value = array == null ? values[i] : array[r + array.GetLowerBound(0), c + array.GetLowerBound(1)];
                        matrix[row++, 3] = value is ExcelMissing || value is ExcelEmpty ? null : value;
                    }
                }
                // Generation and content digest enter GKD's normal fingerprint, including
                // for edits whose UI version did not increase and A->B->A changes.
                string prompt = definition.PromptTemplate + "\n\n自定义公式：" + definition.ExcelName + "；单位：" + definition.OutputUnit
                    + "\n输入参数按附表中的参数名、行、列读取。\n定义标识：" + session + ":" + revision.ToString(CultureInfo.InvariantCulture) + ":" + Digest(definition);
                if (prompt.Length > 16000) return Prefix + "AI 公式提示词与参数说明过长，请缩短定义。";
                return GkdAiFunctions.Gkd(prompt, matrix);
            }
            catch (ArgumentException ex) { return Prefix + ex.Message; }
            catch { return Prefix + "公式输入或定义无效。"; }
        }
        private static string Digest(GkdCustomFormula formula)
        {
            using (var sha = SHA256.Create()) return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(GkdCustomFormulaJson.Export(new[] { formula }))));
        }
        private static long ReferenceCells(ExcelReference reference)
        {
            var areas = reference.InnerReferences;
            if (areas == null || areas.Count == 0) areas = new List<ExcelReference> { reference };
            long actual = 0, rows = 0, columns = 0;
            foreach (var area in areas)
            {
                long h = (long)area.RowLast - area.RowFirst + 1, w = (long)area.ColumnLast - area.ColumnFirst + 1;
                if (h < 1 || w < 1 || h > 10000 || w > 10000 || h * w > 10000 || actual + h * w > 10000) throw new ArgumentException("引用过大，请缩小选择范围。");
                actual += h * w; rows += h; columns = Math.Max(columns, w);
            }
            return rows * columns > 10000 ? actual : rows * columns;
        }
        private static HashSet<string> BuildReservedNames()
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "GKD", "GKD_CUSTOM", "R", "C", "TRUE", "FALSE", "LET", "LAMBDA", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "XLOOKUP", "XMATCH", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER", "VSTACK", "HSTACK", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS", "WRAPROWS", "WRAPCOLS", "TOROW", "TOCOL", "EXPAND", "GROUPBY", "PIVOTBY", "REGEXTEST", "REGEXEXTRACT", "REGEXREPLACE", "PY", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY", "ISOMITTED" };
            foreach (var field in typeof(XlCall).GetFields(BindingFlags.Public | BindingFlags.Static))
                if (field.Name.StartsWith("xlf", StringComparison.Ordinal)) names.Add(field.Name.Substring(3));
            foreach (var type in typeof(GkdCustomFormulaFunctions).Assembly.GetTypes())
                foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
                {
                    var attribute = method.GetCustomAttributes(typeof(ExcelFunctionAttribute), false).FirstOrDefault() as ExcelFunctionAttribute;
                    if (attribute != null) names.Add(string.IsNullOrEmpty(attribute.Name) ? method.Name : attribute.Name);
                    var command = method.GetCustomAttributes(typeof(ExcelCommandAttribute), false).FirstOrDefault() as ExcelCommandAttribute;
                    if (command != null) names.Add(string.IsNullOrEmpty(command.Name) ? method.Name : command.Name);
                }
            return names;
        }
        private static bool IsCellName(string name)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(name, @"\AR[1-9][0-9]*C[1-9][0-9]*\z", System.Text.RegularExpressions.RegexOptions.CultureInvariant)) return true;
            int i = 0, column = 0;
            while (i < name.Length && name[i] >= 'A' && name[i] <= 'Z') { if (i >= 3) return false; column = column * 26 + name[i++] - 'A' + 1; }
            int row;
            return i > 0 && column <= 16384 && i < name.Length && name[i] != '0' && int.TryParse(name.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out row) && row >= 1 && row <= 1048576;
        }
    }
}
