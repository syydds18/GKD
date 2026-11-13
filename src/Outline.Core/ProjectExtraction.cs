using System;
using System.Collections.Generic;

namespace Outline.Core
{
    public enum ProjectType
    {
        Hydropower
    }

    public sealed class ProjectSheetPlan
    {
        public int SheetIndex { get; internal set; }
        public int RowScanStart { get; internal set; }
        public int DeleteColumnsFrom { get; internal set; }
        public bool RatioDriven { get; internal set; }
        public bool DeleteRowsAfterBlankInColumnC { get; internal set; }
        public int DeleteRowCount { get; internal set; }
    }

    public static class ProjectExtraction
    {
        public static ProjectSheetPlan Plan(ProjectType type, int sheetIndex, int maxRows, int maxColumns)
        {
            if (sheetIndex < 1) throw new ArgumentOutOfRangeException("sheetIndex");
            if (maxRows < 1 || maxRows > 1048576) throw new InvalidOperationException("Excel行数超出范围。");
            if (maxColumns < 1 || maxColumns > 16384) throw new InvalidOperationException("Excel列数超出范围。");

            var plan = new ProjectSheetPlan
            {
                SheetIndex = sheetIndex,
                RowScanStart = 0,
                DeleteColumnsFrom = 0,
                RatioDriven = false,
                DeleteRowsAfterBlankInColumnC = false
            };

            if (type != ProjectType.Hydropower) throw new ArgumentOutOfRangeException("type");
            return plan;
        }
    }
}
