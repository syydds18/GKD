using System;
using System.Globalization;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using ExcelDna.Integration;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Outline.Addin
{
    // Entry points are added below; COM work is always deferred to an Excel macro.
    internal static class GkdAiTableFormatting
    {
        private static bool Running = true;
        private const string MixedNotice = "GKD：中英混排已依次设置中文字体和 Times New Roman；中文字体保留效果随 Excel 版本而异，请在本机核验。";
        private static Timer NoticeTimer;
        private static object PreviousStatus;
        private static long NoticeVersion;
        private const int MaxRequests = 256, MaxAttempts = 16;
        // Invoke synchronously from AutoOpen/AutoClose, on Excel's main thread.
        internal static void Start() { Running = true; }
        internal static void Stop()
        {
            Running = false;
            RestoreNotice();
            foreach (Request request in RequestOrder) StopTimer(request);
            Requests.Clear(); RequestOrder.Clear();
        }
        private static readonly Dictionary<string, Request> Requests = new Dictionary<string, Request>(StringComparer.Ordinal);
        private static readonly Queue<Request> RequestOrder = new Queue<Request>();
        private sealed class Request
        {
            internal string Key, Formula, Fingerprint;
            internal ExcelReference Caller;
            internal object[,] Table;
            internal string Text;
            internal int Attempts;
            internal Timer Timer;
        }

        // Call ONLY on Excel's UDF/main thread after the async utility returns a final
        // array. No COM objects are read or captured here. Failed scheduling cannot
        // turn a successful AI result into a UDF error.
        internal static void Schedule(ExcelReference caller, object[,] data)
        {
            if (caller == null || !ValidLayout(data)) return;
            Enqueue(caller, (object[,])data.Clone(), null);
        }
        internal static void ScheduleScalar(ExcelReference caller, string expectedText)
        {
            if (caller == null || expectedText == null || expectedText.StartsWith("GKD：", StringComparison.Ordinal)) return;
            Enqueue(caller, null, expectedText);
        }
        private static void Enqueue(ExcelReference caller, object[,] table, string text)
        {
            if (!Running) return;
            try
            {
                // GET.CELL(41) returns the invariant formula. Re-read it via the same
                // C API in every macro, avoiding Formula2/Formula dialect differences.
                string formula = XlCall.Excel(XlCall.xlfGetCell, 41, caller) as string;
                if (!IsGkdFormula(formula)) return;
                string key = caller.SheetId.ToInt64().ToString(CultureInfo.InvariantCulture) + ":" + caller.RowFirst + ":" + caller.ColumnFirst;
                string fingerprint = Fingerprint(formula, table, text);
                Request previous;
                if (Requests.TryGetValue(key, out previous))
                {
                    if (previous.Fingerprint == fingerprint) return;
                    StopTimer(previous);
                }
                var request = new Request { Key = key, Formula = formula, Fingerprint = fingerprint, Caller = caller, Table = table, Text = text };
                Requests[key] = request;
                RequestOrder.Enqueue(request);
                while (RequestOrder.Count > MaxRequests)
                {
                    Request old = RequestOrder.Dequeue();
                    StopTimer(old);
                    Request current;
                    if (Requests.TryGetValue(old.Key, out current) && object.ReferenceEquals(old, current)) Requests.Remove(old.Key);
                }
                ExcelAsyncUtil.QueueAsMacro(() => Attempt(request));
            }
            catch { /* Formatting is optional side work; preserve the actual result. */ }
        }
        private static void Attempt(Request request)
        {
            Request current;
            if (!Running || !Requests.TryGetValue(request.Key, out current) || !object.ReferenceEquals(current, request)) return;
            request.Attempts++;
            try
            {
                string formula = XlCall.Excel(XlCall.xlfGetCell, 41, request.Caller) as string;
                if (formula != request.Formula) { Forget(request); return; }
                object anchor = ResolveAnchor(request.Caller);
                string nativeFormula = ReadFormula(anchor);
                // OwnsFormula also independently rejects constants/wrapped calls.
                FormatOutcome result = request.Table != null
                    ? TryFormatTable(anchor, nativeFormula, request.Table)
                    : TryFormatScalar(anchor, nativeFormula, request.Text);
                if (result == FormatOutcome.Applied)
                {
                    bool mixed = request.Table == null ? IsMixedText(request.Text) : ContainsMixedText(request.Table);
                    if (mixed) ShowMixedNotice();
                    return;
                }
                if (result == FormatOutcome.Rejected) { Forget(request); return; }
            }
            catch { /* Excel may still be publishing the RTD/dynamic spill. */ }
            if (request.Attempts >= MaxAttempts) { Forget(request); return; }
            // WinForms timer is created in the macro on Excel's UI thread. Tick only
            // queues another macro; it never accesses workbook COM directly.
            request.Timer = new Timer { Interval = 250 };
            request.Timer.Tick += (sender, args) =>
            {
                StopTimer(request);
                try { ExcelAsyncUtil.QueueAsMacro(() => Attempt(request)); }
                catch { Forget(request); }
            };
            request.Timer.Start();
        }
        private static object ResolveAnchor(ExcelReference caller)
        {
            string sheet = XlCall.Excel(XlCall.xlSheetNm, caller) as string;
            if (string.IsNullOrEmpty(sheet)) throw new InvalidOperationException("原工作表已不可用。");
            if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'') sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
            int left = sheet.LastIndexOf('['), right = sheet.LastIndexOf(']');
            if (left < 0 || right <= left + 1 || right == sheet.Length - 1) throw new InvalidOperationException("无法确定原工作簿。");
            dynamic app = ExcelDnaUtil.Application;
            dynamic workbook = app.Workbooks[sheet.Substring(left + 1, right - left - 1)];
            dynamic worksheet = workbook.Worksheets[sheet.Substring(right + 1)];
            return worksheet.Cells[caller.RowFirst + 1, caller.ColumnFirst + 1];
        }
        private static void StopTimer(Request request)
        {
            if (request.Timer == null) return;
            request.Timer.Stop(); request.Timer.Dispose(); request.Timer = null;
        }
        private static void Forget(Request request)
        {
            StopTimer(request);
            Request current;
            if (Requests.TryGetValue(request.Key, out current) && object.ReferenceEquals(current, request)) Requests.Remove(request.Key);
        }
        private static string Fingerprint(string formula, object[,] table, string text)
        {
            var buffer = new StringBuilder();
            Append(buffer, formula);
            if (table == null) { buffer.Append("scalar|"); Append(buffer, text); }
            else
            {
                buffer.Append(table.GetLength(0)).Append(':').Append(table.GetLength(1)).Append('|');
                foreach (object value in table) { Append(buffer, value == null ? "null" : value.GetType().FullName); Append(buffer, Convert.ToString(value, CultureInfo.InvariantCulture)); }
            }
            using (var sha = SHA256.Create()) return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(buffer.ToString())));
        }
        private static void Append(StringBuilder buffer, string value)
        { value = value ?? ""; buffer.Append(value.Length).Append(':').Append(value).Append('|'); }

        internal enum FormatOutcome { Applied, NotReady, Rejected }

        private static void ShowMixedNotice()
        {
            dynamic app = ExcelDnaUtil.Application;
            try
            {
                RestoreNotice();
                PreviousStatus = app.StatusBar;
                app.StatusBar = MixedNotice;
                long version = ++NoticeVersion;
                NoticeTimer = new Timer { Interval = 5000 };
                NoticeTimer.Tick += (sender, args) =>
                {
                    if (version != NoticeVersion) return;
                    NoticeTimer.Stop(); NoticeTimer.Dispose(); NoticeTimer = null;
                    try
                    {
                        ExcelAsyncUtil.QueueAsMacro(() => { if (Running && version == NoticeVersion) RestoreNotice(); });
                    }
                    catch { /* AutoClose also restores our status when still owned. */ }
                };
                NoticeTimer.Start();
            }
            catch { NoticeTimer = null; }
        }
        private static void RestoreNotice()
        {
            if (NoticeTimer != null) { NoticeTimer.Stop(); NoticeTimer.Dispose(); NoticeTimer = null; }
            if (PreviousStatus == null) return;
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                if (object.Equals(app.StatusBar, MixedNotice)) app.StatusBar = PreviousStatus;
            }
            catch { }
            PreviousStatus = null;
        }
        private static bool ContainsMixedText(object[,] table)
        {
            foreach (object value in table) if (IsMixedText(Convert.ToString(value, CultureInfo.InvariantCulture))) return true;
            return false;
        }
        private static bool IsMixedText(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            bool western = false, eastAsian = false;
            foreach (char c in text)
            {
                if (c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z' || c >= '0' && c <= '9') western = true;
                else if (c >= '\u3000' || c >= '\u2e80' && c <= '\u2fff') eastAsian = true;
            }
            return western && eastAsian;
        }

        internal static FormatOutcome TryFormatTable(object target, string expectedFormula, object[,] expected)
        {
            dynamic anchor = target;
            if (!ValidLayout(expected) || !OwnsFormula(target, expectedFormula)) return FormatOutcome.Rejected;
            if (!Convert.ToBoolean(anchor.HasSpill, CultureInfo.InvariantCulture)) return FormatOutcome.NotReady;
            dynamic spill = anchor.SpillingToRange;
            dynamic parent = spill.SpillParent;
            if ((int)spill.Row != (int)anchor.Row || (int)spill.Column != (int)anchor.Column ||
                (int)parent.Row != (int)anchor.Row || (int)parent.Column != (int)anchor.Column ||
                !object.Equals((object)spill.Worksheet, (object)anchor.Worksheet) ||
                !object.Equals((object)parent.Worksheet, (object)anchor.Worksheet)) return FormatOutcome.Rejected;
            if ((int)spill.Rows.Count != expected.GetLength(0) || (int)spill.Columns.Count != expected.GetLength(1)) return FormatOutcome.NotReady;
            object merged = spill.MergeCells;
            if (!(merged is bool) || (bool)merged) return FormatOutcome.Rejected;
            if (!Matches((object)spill.Value2, expected)) return FormatOutcome.NotReady;
            ApplyFonts((object)spill, "宋体", 10.5);
            spill.Font.Bold = false;
            dynamic sheet = anchor.Worksheet;
            int row = anchor.Row, col = anchor.Column, rows = expected.GetLength(0), cols = expected.GetLength(1);
            dynamic title = sheet.Range[sheet.Cells[row, col], sheet.Cells[row, col + cols - 1]];
            title.HorizontalAlignment = 7; // xlCenterAcrossSelection; never merge a dynamic spill.
            dynamic top = sheet.Range[sheet.Cells[row, col], sheet.Cells[row + 1, col + cols - 1]];
            dynamic body = sheet.Range[sheet.Cells[row + 2, col], sheet.Cells[row + rows - 1, col + cols - 1]];
            top.Borders.LineStyle = -4142; // xlLineStyleNone
            body.Borders.LineStyle = 1; // xlContinuous
            body.Borders.Weight = 2; // xlThin
            body.Borders.Color = 0;
            body.Borders[5].LineStyle = -4142; // xlDiagonalDown
            body.Borders[6].LineStyle = -4142; // xlDiagonalUp
            return FormatOutcome.Applied;
        }

        // Formula results do not support Characters rich text. The requested whole-cell
        // Chinese-then-Western sequence is separate, version-dependent behavior.
        internal static FormatOutcome TryFormatScalar(object target, string expectedFormula, string expectedText)
        {
            if (expectedText == null || expectedText.StartsWith("GKD：", StringComparison.Ordinal) || !OwnsFormula(target, expectedFormula)) return FormatOutcome.Rejected;
            dynamic anchor = target;
            object spill;
            try { spill = anchor.HasSpill; }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.ErrorCode == unchecked((int)0x80020003) || ex.ErrorCode == unchecked((int)0x80020006) || ex.ErrorCode == unchecked((int)0x800A01B6))
            { spill = false; } // Member unsupported on pre-dynamic-array Excel; busy/other COM errors still propagate.
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { spill = false; }
            if (!(spill is bool) || (bool)spill) return FormatOutcome.Rejected;
            if (!object.Equals((object)anchor.Value2, expectedText)) return FormatOutcome.NotReady;
            ApplyFonts(target, "宋体", 10.5);
            return FormatOutcome.Applied;
        }

        internal static bool ValidLayout(object[,] expected)
        {
            if (expected == null || expected.GetLowerBound(0) != 0 || expected.GetLowerBound(1) != 0 ||
                expected.GetLength(0) < 3 || expected.GetLength(1) < 1 || expected.Length > 10000 ||
                !(expected[0, 0] is string) || string.IsNullOrWhiteSpace((string)expected[0, 0])) return false;
            for (int c = 0; c < expected.GetLength(1); c++)
            {
                if (c > 0 && !Blank(expected[0, c])) return false;
                if (!Blank(expected[1, c]) || Blank(expected[2, c])) return false;
            }
            return true;
        }

        private static bool Blank(object value) { return value == null || value is string && (string)value == ""; }
        private static bool Matches(object actual, object[,] expected)
        {
            var array = actual as Array;
            if (array == null || array.Rank != 2 || array.GetLength(0) != expected.GetLength(0) || array.GetLength(1) != expected.GetLength(1)) return false;
            for (int r = 0; r < expected.GetLength(0); r++)
                for (int c = 0; c < expected.GetLength(1); c++)
                    if (!SameValue(array.GetValue(r + array.GetLowerBound(0), c + array.GetLowerBound(1)), expected[r, c])) return false;
            return true;
        }
        private static bool SameValue(object actual, object expected)
        {
            if (Blank(actual) && Blank(expected)) return true;
            if (actual == null || expected == null) return false;
            if (actual is string || expected is string || actual is bool || expected is bool) return object.Equals(actual, expected);
            try { return Convert.ToDecimal(actual, CultureInfo.InvariantCulture) == Convert.ToDecimal(expected, CultureInfo.InvariantCulture); }
            catch { return object.Equals(actual, expected); }
        }
        private static void ApplyFonts(object target, string chineseFace, double size)
        {
            dynamic range = target;
            range.Font.Name = chineseFace;
            range.Font.Name = "Times New Roman";
            range.Font.Size = size;
        }
        private static bool OwnsFormula(object target, string expectedFormula)
        {
            dynamic anchor = target;
            if (!Convert.ToBoolean(anchor.HasFormula, CultureInfo.InvariantCulture)) return false;
            string formula = ReadFormula(target);
            return formula == expectedFormula && IsGkdFormula(formula);
        }
        private static string ReadFormula(object target)
        {
            dynamic anchor = target;
            try { return Convert.ToString(anchor.Formula2, CultureInfo.InvariantCulture); }
            catch { return Convert.ToString(anchor.Formula, CultureInfo.InvariantCulture); }
        }
        internal static bool IsGkdFormula(string formula)
        {
            if (formula == null) return false;
            var head = Regex.Match(formula, @"^\s*=\s*@?\s*(?:_xll\.)?GKD\s*\(", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!head.Success) return false;
            int depth = 1; char quoted = '\0';
            for (int i = head.Length; i < formula.Length; i++)
            {
                char c = formula[i];
                if (quoted != '\0') { if (c == quoted) { if (i + 1 < formula.Length && formula[i + 1] == quoted) i++; else quoted = '\0'; } continue; }
                if (c == '"' || c == '\'') quoted = c;
                else if (c == '(') depth++;
                else if (c == ')' && --depth == 0) return string.IsNullOrWhiteSpace(formula.Substring(i + 1));
            }
            return false;
        }
    }
}
