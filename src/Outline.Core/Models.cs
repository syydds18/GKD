using System.Collections.Generic;

namespace Outline.Core
{
    // Row/column keys are 1-based. Levels are 0..7 (Excel Row.OutlineLevel minus one).
    public sealed class SheetData
    {
        public Dictionary<int, Dictionary<int, object>> Values { get; set; } = new Dictionary<int, Dictionary<int, object>>();
        public Dictionary<int, Dictionary<int, string>> Formulas { get; set; } = new Dictionary<int, Dictionary<int, string>>();
        public int MaxRow { get; set; }
        public int MaxCol { get; set; }
        public Dictionary<int, int> OutlineLevels { get; set; } = new Dictionary<int, int>();
        public int OutlineMax { get; set; }
    }
    public sealed class Header
    {
        public int HeaderRow { get; set; }
        public int CodeCol { get; set; }
        public int NameCol { get; set; }
        public bool LowConfidence { get; set; }
    }
    public sealed class Detection
    {
        public int HeaderRow { get; set; }
        public int CodeCol { get; set; }
        public int NameCol { get; set; }
        public Dictionary<int, int> Levels { get; set; } = new Dictionary<int, int>();
        public int MaxLevel { get; set; }
        public bool PartIsFamily { get; set; }
        public List<string> Families { get; set; } = new List<string>();
        public List<string> FamilyList { get; set; } = new List<string>();
        public int LastDataRow { get; set; }
        public int StructuredRows { get; set; }
    }
    public sealed class CodeInfo
    {
        public string Family { get; set; }
        public int Depth { get; set; }
        public string Text { get; set; }
        public List<int> Parts { get; set; } = new List<int>();
    }
    public sealed class NumberValue
    {
        public bool Ok { get; set; }
        public double Value { get; set; }
        public string Error { get; set; }
    }
    public sealed class ValidationError
    {
        public string Kind { get; set; }
        public int Row { get; set; }
        public int Col { get; set; }
        public string Letter { get; set; }
        public string RowName { get; set; }
        public string ColLabel { get; set; }
        public double? ParentValue { get; set; }
        public double? ChildSum { get; set; }
        public double Diff { get; set; }
        public int ChildCount { get; set; }
        public string Basis { get; set; }
        public string Hint { get; set; }
        public int Level { get; set; }
        public bool HasFormula { get; set; }
        public string Formula { get; set; }
        public string Message { get; set; }
    }
    public sealed class ErrorGroup
    {
        public int Row { get; set; }
        public string RowName { get; set; }
        public int Level { get; set; }
        public List<ValidationError> Cells { get; set; } = new List<ValidationError>();
    }
    public sealed class ColumnInfo
    {
        public int Col { get; set; }
        public string Letter { get; set; }
        public string Label { get; set; }
    }
    public sealed class ValidationResult
    {
        public string Status { get; set; } = "na";
        public string Reason { get; set; } = "没有可校验的数据";
        public string LevelSource { get; set; } = "none";
        public int MaxLevel { get; set; }
        public int CheckedCells { get; set; }
        public int CheckedRows { get; set; }
        public int ErrorCount { get; set; }
        public List<ValidationError> Errors { get; set; } = new List<ValidationError>();
        public List<ErrorGroup> ErrorGroups { get; set; } = new List<ErrorGroup>();
        public int ErrorRows { get; set; }
        public double MaxDiff { get; set; }
        public List<ColumnInfo> Columns { get; set; } = new List<ColumnInfo>();
        public bool Truncated { get; set; }
        public int HiddenErrors { get; set; }
        public bool Fallback { get; set; }
    }
}
