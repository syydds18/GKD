using System;
using System.Collections.Generic;
using System.Globalization;

namespace Outline.Addin
{
    // Range copy/paste does not carry native worksheet outlines. Call on a fresh sheet,
    // after formatting/AutoFit, so the final Hidden assignments remain authoritative.
    internal static class HydroOutlineTransfer
    {
        internal static void Copy(dynamic sourceRange, dynamic destinationSheet)
        {
            if ((object)sourceRange == null) throw new ArgumentNullException(nameof(sourceRange));
            if ((object)destinationSheet == null) throw new ArgumentNullException(nameof(destinationSheet));
            if (Convert.ToInt32(sourceRange.Areas.Count) != 1)
                throw new ArgumentException("分级复制仅支持单个连续矩形范围。", nameof(sourceRange));
            object sourceSheet = sourceRange.Parent;
            if (ReferenceEquals(sourceSheet, (object)destinationSheet) || sourceSheet.Equals((object)destinationSheet))
                throw new ArgumentException("目标必须是不同的新工作表，不能修改源工作表。", nameof(destinationSheet));
            int firstRow = Convert.ToInt32(sourceRange.Row);
            int firstColumn = Convert.ToInt32(sourceRange.Column);
            int rows = Convert.ToInt32(sourceRange.Rows.Count);
            int columns = Convert.ToInt32(sourceRange.Columns.Count);
            ValidateExtent(firstRow, rows, 1048576);
            ValidateExtent(firstColumn, columns, 16384);
            int summaryRow = Convert.ToInt32(((dynamic)sourceSheet).Outline.SummaryRow);
            int summaryColumn = Convert.ToInt32(((dynamic)sourceSheet).Outline.SummaryColumn);
            if ((summaryRow != 0 && summaryRow != 1) || (summaryColumn != -4131 && summaryColumn != -4152))
                throw new InvalidOperationException("源工作表汇总行/列方向无效。");
            // Capture and validate BOTH axes before any destination writes. Do not
            // ClearOutline: accidentally passing the source must never destroy it.
            AxisSnapshot rowState = Capture(sourceSheet, false, firstRow, rows);
            AxisSnapshot columnState = Capture(sourceSheet, true, firstColumn, columns);
            CaptureDetails(sourceSheet, false, firstRow, rows, summaryRow == 0, rowState);
            CaptureDetails(sourceSheet, true, firstColumn, columns, summaryColumn == -4131, columnState);
            if (Level((object)Axis((object)destinationSheet, false, 1, rows).OutlineLevel) != 1 ||
                Level((object)Axis((object)destinationSheet, true, 1, columns).OutlineLevel) != 1)
                throw new InvalidOperationException("目标工作表导出区域已有分级，请使用新工作表。");
            destinationSheet.Outline.SummaryRow = summaryRow;
            destinationSheet.Outline.SummaryColumn = summaryColumn;
            Apply((object)destinationSheet, false, rowState);
            Apply((object)destinationSheet, true, columnState);
        }

        private static AxisSnapshot Capture(object sheet, bool columns, int first, int count)
        {
            var snapshot = new AxisSnapshot();
            ReadSegments(first, first + count - 1, first,
                (a, b) => (object)Axis(sheet, columns, a, b).OutlineLevel, Level, snapshot.Levels);
            ReadSegments(first, first + count - 1, first,
                (a, b) => (object)Axis(sheet, columns, a, b).Hidden, Hidden, snapshot.Hidden);
            return snapshot;
        }

        // Excel reports mixed OutlineLevel/Hidden as null or DBNull. Bisect only
        // mixed ranges; a uniform million-row axis costs one property read.
        private static void ReadSegments<T>(int first, int last, int origin,
            Func<int, int, object> read, Func<object, T?> convert, List<Span<T>> spans) where T : struct
        {
            T? state = convert(read(first, last));
            if (state.HasValue)
            {
                int start = first - origin + 1, end = last - origin + 1;
                if (spans.Count > 0 && EqualityComparer<T>.Default.Equals(spans[spans.Count - 1].Value, state.Value))
                    spans[spans.Count - 1].Last = end;
                else spans.Add(new Span<T>(start, end, state.Value));
                return;
            }
            if (first == last) throw new InvalidOperationException("无法读取单行或单列的分级/隐藏状态。");
            int middle = first + (last - first) / 2;
            ReadSegments(first, middle, origin, read, convert, spans);
            ReadSegments(middle + 1, last, origin, read, convert, spans);
        }

        private static void ValidateExtent(int first, int count, int limit)
        {
            if (first < 1 || count < 1 || first > limit || count > limit - first + 1)
                throw new ArgumentException("导出范围超过 Excel 行列限制或为空。");
        }

        private static int? Level(object value)
        {
            if (value == null || value == DBNull.Value) return null;
            double level = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (double.IsNaN(level) || level < 1 || level > 8 || level != Math.Truncate(level))
                throw new InvalidOperationException("Excel 原生分级必须为 1 至 8 的整数。");
            return (int)level;
        }

        private static bool? Hidden(object value)
        {
            return value == null || value == DBNull.Value ? (bool?)null : Convert.ToBoolean(value, CultureInfo.InvariantCulture);
        }

        private static IEnumerable<Span<int>> Groups(AxisSnapshot snapshot)
        {
            // Broad groups first, then nested groups: Group increments each member's
            // level. Repeated identical spans preserve ancestors clipped by export.
            for (int level = 2; level <= 8; level++)
            {
                int first = 0, last = 0;
                foreach (Span<int> span in snapshot.Levels)
                {
                    if (span.Value >= level)
                    {
                        if (first == 0) first = span.First;
                        last = span.Last;
                    }
                    else if (first != 0)
                    {
                        yield return new Span<int>(first, last, level);
                        first = 0;
                    }
                }
                if (first != 0) yield return new Span<int>(first, last, level);
            }
        }

        private static void CaptureDetails(object sourceSheet, bool columns, int origin, int count,
            bool summaryBefore, AxisSnapshot snapshot)
        {
            var seen = new HashSet<int>();
            foreach (Span<int> group in Groups(snapshot))
            {
                int summary = summaryBefore ? group.First - 1 : group.Last + 1;
                // A clipped ancestor may have no summary within the export. Never
                // manufacture one on a detail row, or touch cells outside the clip.
                if (summary < 1 || summary > count || !seen.Add(summary)) continue;
                int sourceIndex = origin + summary - 1;
                try
                {
                    object state = Axis(sourceSheet, columns, sourceIndex, sourceIndex).ShowDetail;
                    if (state is bool) snapshot.Details.Add(new Span<bool>(summary, summary, (bool)state));
                }
                catch (System.Runtime.InteropServices.COMException) { /* Not a valid summary in this Excel version. */ }
            }
        }

        private static void Apply(object sheet, bool columns, AxisSnapshot snapshot)
        {
            foreach (Span<int> group in Groups(snapshot))
                Axis(sheet, columns, group.First, group.Last).Group();
            // Expand first, then collapse inner-to-outer. ShowDetail has visibility
            // side effects; exact source Hidden segments are restored AFTER all of it.
            foreach (Span<bool> summary in snapshot.Details)
                if (summary.Value) SetDetail(sheet, columns, summary);
            for (int i = snapshot.Details.Count - 1; i >= 0; i--)
                if (!snapshot.Details[i].Value) SetDetail(sheet, columns, snapshot.Details[i]);
            foreach (Span<bool> span in snapshot.Hidden)
                Axis(sheet, columns, span.First, span.Last).Hidden = span.Value;
        }

        private static void SetDetail(object sheet, bool columns, Span<bool> summary)
        {
            try { Axis(sheet, columns, summary.First, summary.First).ShowDetail = summary.Value; }
            catch (System.Runtime.InteropServices.COMException) { /* Best effort; exact Hidden states still restored. */ }
        }

        private static dynamic Axis(object sheetObject, bool columns, int first, int last)
        {
            dynamic sheet = sheetObject;
            if (!columns) return sheet.Rows[first.ToString(CultureInfo.InvariantCulture) + ":" + last.ToString(CultureInfo.InvariantCulture)];
            return sheet.Range[sheet.Cells[1, first], sheet.Cells[1, last]].EntireColumn;
        }

        private sealed class AxisSnapshot
        {
            internal readonly List<Span<int>> Levels = new List<Span<int>>();
            internal readonly List<Span<bool>> Hidden = new List<Span<bool>>();
            internal readonly List<Span<bool>> Details = new List<Span<bool>>();
        }

        private sealed class Span<T>
        {
            internal readonly int First;
            internal int Last;
            internal readonly T Value;
            internal Span(int first, int last, T value) { First = first; Last = last; Value = value; }
        }
    }
}
