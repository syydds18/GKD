using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace Outline.Addin
{
    internal static class GkdAiActions
    {
        internal const int CellLimit = 10000;

        // Controller invokes this synchronously from the Ribbon macro on Excel's UI thread.
        internal static void Run(string action)
        {
            try
            {
                RunCore(ExcelDnaUtil.Application, action, message =>
                    MessageBox.Show(message, "搞快点 AI", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes);
            }
            catch (Exception ex)
            {
                MessageBox.Show("AI 操作未完成：" + ex.Message, "搞快点 AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        internal static void RunCore(object application, string action, Func<string, bool> confirm)
        {
            dynamic app = application;
            if (action == "aiCancel")
            {
                GkdAiFunctions.Scheduler.CancelAll();
                app.StatusBar = "GKD：已取消全部排队和运行中的 AI 任务；使用刷新 AI 结果重新运行。";
                return;
            }
            if (action != "aiRefresh" && action != "aiValues") throw new ArgumentException("未知 AI 操作。", "action");
            var anchors = SnapshotAnchors((object)app.Selection);
            if (anchors.Count == 0) { app.StatusBar = "GKD：所选区域中没有 GKD 公式结果。"; return; }
            if (action == "aiValues") { Freeze(application, anchors, confirm); return; }
            // Advance all revisions before any calculation; late async replies retain the old key.
            foreach (var anchor in anchors) GkdAiFunctions.RefreshCell(anchor.Identity);
            foreach (var anchor in anchors)
            {
                dynamic cell = anchor.Cell;
                cell.Dirty();
                cell.Calculate();
            }
            app.StatusBar = "GKD：已刷新 " + anchors.Count + " 个所选 AI 公式。";
        }

        private sealed class Anchor
        {
            internal object Cell;
            internal string Identity;
        }

        private static List<Anchor> SnapshotAnchors(object selected)
        {
            if (selected == null) throw new InvalidOperationException("请先选择单元格区域。");
            dynamic selection = selected;
            long count = Convert.ToInt64(selection.CountLarge, CultureInfo.InvariantCulture);
            if (count < 1 || count > CellLimit) throw new InvalidOperationException("请选择 1 至 10000 个单元格。");
            var result = new List<Anchor>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            int visited = 0;
            foreach (dynamic area in selection.Areas)
                foreach (dynamic selectedCell in area.Cells)
                {
                    if (++visited > CellLimit) throw new InvalidOperationException("选择区域超过 10000 个单元格。");
                    object cell = selectedCell;
                    try { object parent = selectedCell.SpillParent; if (parent != null) cell = parent; }
                    catch (COMException) { } // Non-spill cells and Excel versions without dynamic arrays.
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                    dynamic anchor = cell;
                    if (!Convert.ToBoolean(anchor.HasFormula, CultureInfo.InvariantCulture)) continue;
                    string formula = ReadFormula(cell);
                    if (!IsGkdFormula(formula)) continue;
                    int row = Convert.ToInt32(anchor.Row, CultureInfo.InvariantCulture) - 1;
                    int col = Convert.ToInt32(anchor.Column, CultureInfo.InvariantCulture) - 1;
                    // Selection is on the active sheet, also used by the four-coordinate constructor.
                    var reference = new ExcelReference(row, row, col, col);
                    string identity = GkdAiFunctions.CellIdentity(reference.SheetId, row, col);
                    if (seen.Add(identity)) result.Add(new Anchor { Cell = cell, Identity = identity });
                }
            return result;
        }

        private sealed class FreezeItem
        {
            internal object Range;
            internal string Label;
            internal long Count;
            internal object Anchor;
            internal string Formula;
        }

        private static object ResultRange(object anchor)
        {
            dynamic cell = anchor;
            try { object range = cell.SpillingToRange; return range ?? anchor; }
            catch (COMException) { return anchor; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return anchor; }
        }

        private static void Freeze(object application, List<Anchor> anchors, Func<string, bool> confirm)
        {
            var plan = new List<FreezeItem>();
            long total = 0;
            foreach (var anchor in anchors)
            {
                object range = ResultRange(anchor.Cell);
                dynamic target = range;
                long count = Convert.ToInt64(target.CountLarge, CultureInfo.InvariantCulture);
                total += count;
                if (count < 1 || total > CellLimit) throw new InvalidOperationException("完整结果范围超过 10000 个单元格。");
                ValidateFreezeRange(range);
                plan.Add(new FreezeItem { Range = range, Count = count, Label = Convert.ToString(target.Address, CultureInfo.InvariantCulture), Anchor = anchor.Cell, Formula = ReadFormula(anchor.Cell) });
            }
            var labels = new List<string>();
            foreach (var item in plan) labels.Add(item.Label + "（" + item.Count + " 格）");
            if (!confirm("将 " + anchors.Count + " 个 GKD 结果、共 " + total + " 个单元格转为值：\n" +
                string.Join("\n", labels) + "\n即使仅选择溢出子单元格，也会转换完整溢出范围。公式将被移除，不能自动恢复。是否继续？")) return;
            // A modal confirmation can pump messages. Recheck every target before any copy/write.
            foreach (var item in plan)
            {
                dynamic current = ResultRange(item.Anchor);
                if (!string.Equals(ReadFormula(item.Anchor), item.Formula, StringComparison.Ordinal) ||
                    !string.Equals(Convert.ToString(current.Address, CultureInfo.InvariantCulture), item.Label, StringComparison.Ordinal) ||
                    Convert.ToInt64(current.CountLarge, CultureInfo.InvariantCulture) != item.Count)
                    throw new InvalidOperationException("确认期间 GKD 公式或溢出范围已变化，请重新选择后操作。");
                ValidateFreezeRange((object)current);
            }
            dynamic app = application;
            bool events = Convert.ToBoolean(app.EnableEvents, CultureInfo.InvariantCulture);
            try
            {
                app.EnableEvents = false;
                foreach (var item in plan)
                {
                    dynamic target = item.Range;
                    // Excel captures the entire spill on its native clipboard BEFORE removing
                    // the anchor formula; no Value2 assignment can coerce strings or errors.
                    target.Copy();
                    target.PasteSpecial(-4163, Type.Missing, Type.Missing, Type.Missing); // xlPasteValues
                }
                app.StatusBar = "GKD：已将 " + total + " 个结果单元格转为值。";
            }
            finally
            {
                try { app.CutCopyMode = false; }
                finally { app.EnableEvents = events; }
            }
        }

        private static void ValidateFreezeRange(object range)
        {
            dynamic target = range;
            if (Convert.ToBoolean(target.Worksheet.ProtectContents, CultureInfo.InvariantCulture))
                throw new InvalidOperationException("结果所在工作表已保护，请先取消保护。");
            // Mixed tri-state values are unsafe too; fail closed instead of treating null as false.
            RequireFalse((object)target.MergeCells, "结果范围包含合并单元格。");
            RequireFalse((object)target.HasArray, "不能转换旧式数组公式结果。");
            foreach (dynamic cell in target.Cells)
            {
                RequireFalse((object)cell.MergeCells, "结果范围包含合并单元格。");
                RequireFalse((object)cell.HasArray, "不能转换旧式数组公式结果。");
            }
            object values = target.Value2;
            var matrix = values as Array;
            if (matrix == null) ValidateResult(values);
            else foreach (object value in matrix) ValidateResult(value);
        }

        private static void RequireFalse(object value, string message)
        {
            if (!(value is bool) || (bool)value) throw new InvalidOperationException(message);
        }

        private static void ValidateResult(object value)
        {
            // Excel COM returns VT_ERROR as ErrorWrapper (and some versions as its signed
            // HRESULT Int32). Ordinary numeric worksheet values are doubles, not Int32.
            if (value is ExcelError || value is ErrorWrapper || (value is int && (int)value < 0))
                throw new InvalidOperationException("结果含 Excel 错误或仍在等待 AI（#N/A），请完成后再转为值。");
            var text = value as string;
            if (text != null && Regex.IsMatch(text, @"^\s*GKD\s*[:：]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                throw new InvalidOperationException("结果含 GKD 失败或取消提示，请刷新并完成后再转为值。");
        }

        private static string ReadFormula(object target)
        {
            dynamic cell = target;
            try { return Convert.ToString(cell.Formula2, CultureInfo.InvariantCulture); }
            catch (COMException) { return Convert.ToString(cell.Formula, CultureInfo.InvariantCulture); }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return Convert.ToString(cell.Formula, CultureInfo.InvariantCulture); }
        }

        private static string CustomAliasName(string formula)
        {
            int open = formula.IndexOf('(');
            if (open < 0) return null;
            int index = open + 1;
            while (index < formula.Length && char.IsWhiteSpace(formula[index])) index++;
            if (index >= formula.Length || formula[index++] != '"') return null;
            var name = new System.Text.StringBuilder();
            while (index < formula.Length)
            {
                char ch = formula[index++];
                if (ch != '"') { name.Append(ch); continue; }
                if (index < formula.Length && formula[index] == '"') { name.Append('"'); index++; continue; }
                return name.ToString();
            }
            return null;
        }

        // Only a complete direct call owns the cell's result. Wrappers/suffixes are not ours.
        internal static bool IsGkdFormula(string formula)
        {
            if (formula == null) return false;
            var head = Regex.Match(formula, @"^\s*=\s*@?\s*(?:_xll\.)?([A-Z_][A-Z0-9_]*)\s*\(", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!head.Success) return false;
            string name = head.Groups[1].Value;
            if (string.Equals(name, "GKD", StringComparison.OrdinalIgnoreCase)) { }
            else
            {
                if (string.Equals(name, "GKD_CUSTOM", StringComparison.OrdinalIgnoreCase))
                {
                    name = CustomAliasName(formula);
                    if (name == null) return false;
                }
                if (!GkdCustomFormulaFunctions.IsCustomName(name) || !GkdCustomFormulaFunctions.IsAiFormula(name)) return false;
            }
            int depth = 1;
            bool quoted = false, sheetQuoted = false;
            for (int i = head.Length; i < formula.Length; i++)
            {
                char ch = formula[i];
                if (quoted)
                {
                    if (ch == '"') { if (i + 1 < formula.Length && formula[i + 1] == '"') i++; else quoted = false; }
                    continue;
                }
                if (sheetQuoted)
                {
                    if (ch == '\'') { if (i + 1 < formula.Length && formula[i + 1] == '\'') i++; else sheetQuoted = false; }
                    continue;
                }
                if (ch == '"') quoted = true;
                else if (ch == '\'') sheetQuoted = true;
                else if (ch == '(') depth++;
                else if (ch == ')' && --depth == 0) return string.IsNullOrWhiteSpace(formula.Substring(i + 1));
            }
            return false;
        }
    }
}
