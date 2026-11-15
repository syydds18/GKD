using System;
using System.Collections.Generic;

namespace Outline.Addin
{
    internal static class HydroRangeCopy
    {
        internal static void CopyValuesAndFormats(dynamic source, dynamic target)
        {
            int rows = Convert.ToInt32(source.Rows.Count);
            int columns = Convert.ToInt32(source.Columns.Count);
            List<MergeArea> merges = ReadMerges(source, rows, columns);

            target.UnMerge();
            foreach (MergeArea merge in merges)
                target.Cells[merge.Row, merge.Column].Resize[merge.Rows, merge.Columns].Merge();

            foreach (CopyBlock block in PlanUnmergedBlocks(rows, columns, merges))
            {
                dynamic from = SourceBlock(source, block.SourceRow, block.SourceColumn, block.Rows, block.Columns);
                dynamic to = target.Cells[block.SourceRow, block.SourceColumn].Resize[block.Rows, block.Columns];
                from.Copy();
                to.PasteSpecial(-4122, Type.Missing, Type.Missing, Type.Missing); // xlPasteFormats, with no merges in the block
                to.PasteSpecial(-4163, Type.Missing, Type.Missing, Type.Missing); // xlPasteValues preserves errors
            }

            foreach (MergeArea merge in merges)
            {
                dynamic fromArea = SourceBlock(source, merge.Row, merge.Column, merge.Rows, merge.Columns);
                dynamic toArea = target.Cells[merge.Row, merge.Column].Resize[merge.Rows, merge.Columns];
                // Exactly one matching merge at a time: never paste a heterogeneous merge grid.
                fromArea.Copy();
                toArea.PasteSpecial(-4122, Type.Missing, Type.Missing, Type.Missing);
                toArea.PasteSpecial(-4163, Type.Missing, Type.Missing, Type.Missing);
            }
            dynamic sourceColumns = source.Parent.Cells[Convert.ToInt32(source.Row), Convert.ToInt32(source.Column)].Resize[rows, columns];
            sourceColumns.Copy();
            target.PasteSpecial(8, Type.Missing, Type.Missing, Type.Missing); // xlPasteColumnWidths
        }

        private static dynamic SourceBlock(dynamic source, int row, int column, int rows, int columns)
        {
            return source.Parent.Cells[Convert.ToInt32(source.Row) + row - 1, Convert.ToInt32(source.Column) + column - 1].Resize[rows, columns];
        }

        private static List<MergeArea> ReadMerges(dynamic source, int rows, int columns)
        {
            var result = new List<MergeArea>();
            var seen = new HashSet<MergeArea>();
            object hasMerges = source.MergeCells;
            if (hasMerges is bool && !(bool)hasMerges) return result;
            int sourceRow = Convert.ToInt32(source.Row);
            int sourceColumn = Convert.ToInt32(source.Column);
            for (int row = 1; row <= rows; row++)
            {
                object rowMerges = source.Cells[row, 1].Resize[1, columns].MergeCells;
                if (rowMerges is bool && !(bool)rowMerges) continue;
                for (int column = 1; column <= columns;)
                {
                    dynamic cell = source.Cells[row, column];
                    object merged = cell.MergeCells;
                    if (!(merged is bool) || !(bool)merged) { column++; continue; }

                    dynamic area = cell.MergeArea;
                    int areaRow = Convert.ToInt32(area.Row) - sourceRow + 1;
                    int areaColumn = Convert.ToInt32(area.Column) - sourceColumn + 1;
                    int height = Convert.ToInt32(area.Rows.Count);
                    int width = Convert.ToInt32(area.Columns.Count);
                    if (areaRow < 1 || areaColumn < 1 || areaRow + height - 1 > rows || areaColumn + width - 1 > columns)
                        throw new InvalidOperationException("导出范围不能截断合并单元格，请调整打印范围。");
                    var merge = new MergeArea(areaRow, areaColumn, height, width);
                    if (seen.Add(merge)) result.Add(merge);
                    column = areaColumn + width;
                }
            }
            return result;
        }

        internal static List<CopyBlock> PlanUnmergedBlocks(int rows, int columns, IList<MergeArea> merges)
        {
            var result = new List<CopyBlock>();
            var boundaries = new SortedSet<int> { 1, rows + 1 };
            foreach (var merge in merges)
            {
                boundaries.Add(merge.Row);
                boundaries.Add(merge.Row + merge.Rows);
            }
            var starts = new List<int>(boundaries);
            for (int i = 0; i + 1 < starts.Count; i++)
            {
                int row = starts[i], height = starts[i + 1] - row;
                var active = new List<MergeArea>();
                foreach (var merge in merges)
                    if (row >= merge.Row && row < merge.Row + merge.Rows) active.Add(merge);
                active.Sort((a, b) => a.Column.CompareTo(b.Column));
                int column = 1;
                foreach (var merge in active)
                {
                    if (merge.Column > column) result.Add(new CopyBlock(row, column, height, merge.Column - column));
                    column = Math.Max(column, merge.Column + merge.Columns);
                }
                if (column <= columns) result.Add(new CopyBlock(row, column, height, columns - column + 1));
            }
            return result;
        }

        internal struct MergeArea : IEquatable<MergeArea>
        {
            internal readonly int Row, Column, Rows, Columns;
            internal MergeArea(int row, int column, int rows, int columns) { Row = row; Column = column; Rows = rows; Columns = columns; }
            public bool Equals(MergeArea other) { return Row == other.Row && Column == other.Column && Rows == other.Rows && Columns == other.Columns; }
            public override bool Equals(object value) { return value is MergeArea && Equals((MergeArea)value); }
            public override int GetHashCode() { return (((Row * 397) ^ Column) * 397 ^ Rows) * 397 ^ Columns; }
            public override string ToString() { return Row + "," + Column + ":" + Rows + "x" + Columns; }
        }

        internal struct CopyBlock : IEquatable<CopyBlock>
        {
            internal readonly int SourceRow, SourceColumn, Rows, Columns;
            internal CopyBlock(int row, int column, int rows, int columns) { SourceRow = row; SourceColumn = column; Rows = rows; Columns = columns; }
            public bool Equals(CopyBlock other) { return SourceRow == other.SourceRow && SourceColumn == other.SourceColumn && Rows == other.Rows && Columns == other.Columns; }
            public override bool Equals(object value) { return value is CopyBlock && Equals((CopyBlock)value); }
            public override int GetHashCode() { return (((SourceRow * 397) ^ SourceColumn) * 397 ^ Rows) * 397 ^ Columns; }
            public override string ToString() { return SourceRow + "," + SourceColumn + ":" + Rows + "x" + Columns; }
        }
    }
}
