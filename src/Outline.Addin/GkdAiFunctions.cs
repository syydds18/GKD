using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace Outline.Addin
{
    public static class GkdAiFunctions
    {
        internal static readonly GkdAiScheduler Scheduler = new GkdAiScheduler();
        private static readonly object Gate = new object();
        private static long sequence;
        private static readonly Dictionary<string, long> Revisions = new Dictionary<string, long>();

        [ExcelFunction(Name = "GKD", Category = "搞快点 AI", Description = "单次 AI 处理文本或明确引用的数据；数据可省略；可用 (A1,C3:E3) 指定不连续单元格", IsMacroType = true)]
        public static object Gkd(
            [ExcelArgument(Name = "提示词", Description = "处理规则、问题或存放规则的单元格")] object promptValue,
            [ExcelArgument(Name = "数据", Description = "可省略；单元格、连续区域，或括号包住的不连续引用，如 (A1,C3:E3)", AllowReference = true)] object data = null,
            [ExcelArgument(Name = "输出类型", Description = "可省略，默认单值；填 1 返回二维表格，0 或省略返回单值")] object outputMode = null)
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNA;
            try
            {
                if ((promptValue is ExcelMissing || promptValue == null) && (data is ExcelMissing || data == null) && (outputMode is ExcelMissing || outputMode == null)) return "感谢使用！";
                if (promptValue is ExcelError) return promptValue;
                string prompt = promptValue is ExcelMissing || promptValue is ExcelEmpty || promptValue == null ? "" : Convert.ToString(promptValue, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(prompt)) return "GKD：请输入提示词。";
                string mode = GkdAiProtocol.NormalizeMode(outputMode is ExcelMissing || outputMode is ExcelEmpty || outputMode == null ? "单值" : Convert.ToString(outputMode, CultureInfo.InvariantCulture));
                object[,] values = null;
                // Missing/null are the optional argument defaults; ExcelEmpty is an
                // explicitly supplied blank cell and must not become pure-text mode.
                bool hasData = !(data is ExcelMissing || data == null);
                if (hasData)
                {
                    values = data as object[,];
                    if (data is ExcelReference) values = ReadReference((ExcelReference)data);
                    if (values == null) values = new object[,] { { data } };
                    if (values.Length > 10000) return "GKD：数据区域应为 1 至 10000 个单元格。";
                }
                var snapshot = values == null ? null : new object[values.GetLength(0), values.GetLength(1)];
                if (values != null)
                {
                    bool allEmpty = true;
                    int r0 = values.GetLowerBound(0), c0 = values.GetLowerBound(1);
                    for (int r = 0; r < snapshot.GetLength(0); r++)
                        for (int c = 0; c < snapshot.GetLength(1); c++)
                        {
                            object value = values[r + r0, c + c0];
                            if (value is ExcelError) return value;
                            snapshot[r, c] = value is ExcelEmpty || value is ExcelMissing ? null : value;
                            if (snapshot[r, c] != null && !(value is string && string.IsNullOrWhiteSpace((string)value))) allEmpty = false;
                        }
                    if (allEmpty) return "GKD：对应数据为空，请检查。";
                }
                if (mode == "表格" && !SupportsDynamicArrays()) return "GKD：表格输出需要支持动态数组的 Excel（Microsoft 365 / 2021 / 2024）。";
                GkdAiProfile profile = GkdAiStore.ResolveDefault();
                profile.Validate();
                string input = GkdAiProtocol.BuildPrompt(prompt, snapshot, mode);
                string key = Fingerprint(profile, input);
                string caller = CallerIdentity();
                long revision;
                lock (Gate) { if (!Revisions.TryGetValue(caller, out revision)) Revisions[caller] = revision = 0; }
                string requestKey = key + ":" + revision.ToString(CultureInfo.InvariantCulture);
                object result = ExcelAsyncUtil.RunTaskWithCancellation<object>("GKD", new object[] { caller, requestKey },
                    token => Execute(profile, prompt, snapshot, mode, requestKey, token));
                var reference = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
                if (reference != null)
                {
                    // Scheduling is on the UDF/Excel thread, never in Execute's worker.
                    var table = result as object[,];
                    if (mode == "表格" && table != null) GkdAiTableFormatting.Schedule(reference, table);
                    var text = result as string;
                    if (mode == "单值" && text != null && !text.StartsWith("GKD：", StringComparison.Ordinal))
                        GkdAiTableFormatting.ScheduleScalar(reference, text);
                }
                return result;
            }
            catch (ArgumentException ex) { return "GKD：" + ex.Message; }
            catch { return "GKD：配置或输入无效，请在 AI 设置中检查。"; }
        }

        private static async Task<object> Execute(GkdAiProfile profile, string prompt, object[,] data, string mode, string key, CancellationToken token)
        {
            try
            {
                return await Scheduler.RunAsync(key, async ct =>
                {
                    var result = await GkdAiClient.ExecuteAsync(profile, prompt, data, mode, ct).ConfigureAwait(false);
                    return result.Value;
                }, token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { return "GKD：已取消；使用刷新 AI 结果重新运行。"; }
            catch (System.IO.InvalidDataException ex) { return "GKD：" + ex.Message; }
            catch (TimeoutException ex) { return "GKD：" + ex.Message; }
            catch { return "GKD：请求失败，请检查 AI 设置中的连接与表格测试。"; }
        }

        internal static object[,] ReadReference(ExcelReference reference)
        {
            var areas = reference.InnerReferences;
            if (areas == null || areas.Count == 0) areas = new List<ExcelReference> { reference };
            int rows = 0, columns = 0, actualCells = 0;
            foreach (var area in areas)
            {
                int h = area.RowLast - area.RowFirst + 1, w = area.ColumnLast - area.ColumnFirst + 1;
                long areaCells = (long)h * w;
                if (h < 1 || w < 1 || areaCells > 10000 || actualCells > 10000 - areaCells || rows > 10000 - h)
                    throw new ArgumentException("不连续引用过大，请缩小选择范围。");
                rows += h;
                columns = Math.Max(columns, w);
                actualCells += (int)areaCells;
            }
            if (actualCells == 0) throw new ArgumentException("不连续引用为空，请缩小选择范围。");

            // Different-width areas are normally padded with nulls to make one
            // rectangular request. If that padding alone exceeds the input limit,
            // preserve every referenced cell in area/row/column order as one column.
            bool flatten = (long)rows * columns > 10000;
            var values = flatten ? new object[actualCells, 1] : new object[rows, columns];
            int row = 0, flat = 0;
            foreach (var area in areas)
            {
                object value = area.GetValue();
                var array = value as Array;
                int h = area.RowLast - area.RowFirst + 1, w = area.ColumnLast - area.ColumnFirst + 1;
                if (h == 1 && w == 1)
                {
                    if (flatten) values[flat++, 0] = value;
                    else values[row, 0] = value;
                }
                else if (array != null && array.Rank == 2)
                {
                    for (int y = 0; y < h; y++)
                        for (int x = 0; x < w; x++)
                        {
                            object item = array.GetValue(y + array.GetLowerBound(0), x + array.GetLowerBound(1));
                            if (flatten) values[flat++, 0] = item;
                            else values[row + y, x] = item;
                        }
                }
                else throw new ArgumentException("无法读取不连续引用中的单元格。");
                row += h;
            }
            return values;
        }
        private static bool SupportsDynamicArrays()
        {
            // Same capability probe as Excel-DNA's official dynamic array tutorial.
            try { XlCall.Excel(614, new object[] { 1 }, new object[] { true }); return true; }
            catch { return false; }
        }
        internal static string CallerIdentity()
        {
            var reference = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            if (reference == null) return "noncell";
            return CellIdentity(reference.SheetId, reference.RowFirst, reference.ColumnFirst);
        }
        internal static string CellIdentity(IntPtr sheet, int row, int column)
        { return sheet.ToInt64().ToString(CultureInfo.InvariantCulture) + ":" + row + ":" + column; }
        internal static void ConfigurationChanged()
        {
            // Settings are not worksheet dependencies: preserve completed cells until
            // explicit refresh, but never let an old pending request publish afterward.
            Scheduler.CancelAll();
            Scheduler.ClearCache();
            GkdAiTableFormatting.Stop();
            GkdAiTableFormatting.Start();
        }
        internal static void RefreshCell(string identity)
        {
            lock (Gate)
            {
                Revisions[identity] = ++sequence;
            }
            // Bounded cache; changing revisions prevents an earlier in-flight reply
            // from becoming the refreshed cell's result.
            Scheduler.ClearCache();
        }
        private static string Fingerprint(GkdAiProfile profile, string input)
        {
            var text = new StringBuilder(input);
            foreach (var property in profile.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public).OrderBy(x => x.Name, StringComparer.Ordinal))
            {
                string value = Convert.ToString(property.GetValue(profile, null), CultureInfo.InvariantCulture) ?? "";
                text.Append('|').Append(property.Name).Append(':').Append(value.Length).Append(':').Append(value);
            }
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text.ToString()))).Replace("-", "");
        }
    }
}
