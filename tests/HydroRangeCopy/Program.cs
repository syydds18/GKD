using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using Outline.Addin;

static class Program
{
    static int Main()
    {
        var blocks = HydroRangeCopy.PlanUnmergedBlocks(4, 5, new[]
        {
            new HydroRangeCopy.MergeArea(2, 2, 2, 3),
            new HydroRangeCopy.MergeArea(4, 1, 1, 2)
        });
        var expected = new[]
        {
            new HydroRangeCopy.CopyBlock(1, 1, 1, 5),
            new HydroRangeCopy.CopyBlock(2, 1, 2, 1),
            new HydroRangeCopy.CopyBlock(2, 5, 2, 1),
            new HydroRangeCopy.CopyBlock(4, 3, 1, 3)
        };
        if (!blocks.SequenceEqual(expected)) throw new Exception("unexpected merge-safe blocks: " + string.Join(";", blocks));
        if (!HydroRangeCopy.PlanUnmergedBlocks(20, 8, new HydroRangeCopy.MergeArea[0]).SequenceEqual(new[] { new HydroRangeCopy.CopyBlock(1, 1, 20, 8) })) throw new Exception("unmerged ranges must be batched");
        if (HydroRangeCopy.PlanUnmergedBlocks(4, 5, new[] { new HydroRangeCopy.MergeArea(1, 1, 4, 5) }).Count != 0) throw new Exception("fully merged range has no unmerged cells");
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../../"));
        var dialog = File.ReadAllText(Path.Combine(root, "src/Outline.Addin/HydroExportDialogs.cs"));
        if (!dialog.Contains("footer.Controls.Add(ok);") || !dialog.Contains("footer.Controls.Add(openFile);")) throw new Exception("open-file must be added without removing OK");
        if (!dialog.Contains("Clipboard.SetDataObject(data, true, 10, 100)") || !dialog.Contains("ApartmentState.STA")) throw new Exception("file clipboard must persist and retry on an STA thread");
        var tall = HydroRangeCopy.PlanUnmergedBlocks(10000, 5, new[]{new HydroRangeCopy.MergeArea(1,2,10000,3)});
        if (tall.Count != 2 || tall.Any(x=>x.Rows!=10000)) throw new Exception("adjacent free columns around a tall merge must be copied in two bands, not per row");
        TestTransfer();
        Console.WriteLine("PASS merge-aware planner/transfer and dialog contracts");
        return 0;
    }

    static void TestTransfer()
    {
        var source = new FakeSheet(9, 10);
        source.Merges.Add(new Box(5, 4, 1, 5));
        source.Merges.Add(new Box(6, 5, 2, 2));
        source.Values[5, 4] = "#DIV/0!";
        source.Values[6, 5] = 42.0;
        source.Values[6, 4] = "=literal text, not a formula";
        FakeRange.CopyCalls = 0;
        var target = new FakeSheet(4, 5);
        HydroRangeCopy.CopyValuesAndFormats(new FakeRange(source, new Box(5, 4, 4, 5)), new FakeRange(target, new Box(1, 1, 4, 5)));
        if (!Equals(target.Values[1, 1], "#DIV/0!") || !Equals(target.Values[2, 2], 42.0) || !Equals(target.Values[2, 1], "=literal text, not a formula")) throw new Exception("value transfer changed anchors, errors, or text");
        if (source.Merges.Count != 2 || target.Merges.Count != 2 || !target.WidthsCopied) throw new Exception("merge/width topology lost");
        if (!target.Merges.Contains(new Box(1, 1, 1, 5)) || !target.Merges.Contains(new Box(2, 2, 2, 2))) throw new Exception("merge offsets incorrect");
        var plan=HydroRangeCopy.PlanUnmergedBlocks(4,5,new[]{new HydroRangeCopy.MergeArea(1,1,1,5),new HydroRangeCopy.MergeArea(2,2,2,2)});
        if (FakeRange.CopyCalls != plan.Count + 3) throw new Exception("copy once per block/merge, reuse clipboard for format and value paste");
        try
        {
            HydroRangeCopy.CopyValuesAndFormats(new FakeRange(source, new Box(5, 5, 4, 4)), new FakeRange(new FakeSheet(4, 4), new Box(1, 1, 4, 4)));
            throw new Exception("partial merges must fail before writing");
        }
        catch (InvalidOperationException ex) { if (!ex.Message.Contains("截断")) throw; }
    }
}

public struct Box : IEquatable<Box>
{
    public readonly int Row, Column, Rows, Columns;
    public Box(int r, int c, int h, int w) { Row=r; Column=c; Rows=h; Columns=w; }
    public bool Contains(int r, int c) { return r>=Row && c>=Column && r<Row+Rows && c<Column+Columns; }
    public bool Equals(Box b) { return Row==b.Row && Column==b.Column && Rows==b.Rows && Columns==b.Columns; }
    public override bool Equals(object o) { return o is Box && Equals((Box)o); }
    public override int GetHashCode() { return Row ^ Column ^ Rows ^ Columns; }
}
public sealed class FakeSheet
{
    public readonly List<Box> Merges = new List<Box>();
    public readonly object[,] Values;
    public bool WidthsCopied;
    public FakeSheet(int rows, int cols) { Values=new object[rows+1,cols+1]; }
    public FakeCells Cells { get { return new FakeCells(this, 1, 1); } }
}
public sealed class FakeCells
{
    readonly FakeSheet sheet; readonly int row, column;
    public FakeCells(FakeSheet s, int r, int c) { sheet=s; row=r; column=c; }
    public FakeRange this[int r, int c] { get { return new FakeRange(sheet, new Box(row+r-1,column+c-1,1,1)); } }
}
public sealed class FakeResize
{
    readonly FakeRange range;
    public FakeResize(FakeRange r) { range=r; }
    public FakeRange this[int rows, int cols] { get { return new FakeRange(range.Parent, new Box(range.Row,range.Column,rows,cols)); } }
}
public sealed class FakeCount
{
    public int Count { get; }
    public FakeCount(int n) { Count=n; }
}
public sealed class FakeRange
{
    static FakeRange clipboard;
    readonly Box box;
    public FakeSheet Parent { get; }
    public FakeRange(FakeSheet s, Box b) { Parent=s; box=b; }
    public int Row { get { return box.Row; } }
    public int Column { get { return box.Column; } }
    public FakeCount Rows { get { return new FakeCount(box.Rows); } }
    public FakeCount Columns { get { return new FakeCount(box.Columns); } }
    public FakeResize Resize { get { return new FakeResize(this); } }
    public FakeCells Cells { get { return new FakeCells(Parent,Row,Column); } }
    public object Value2 { get { throw new Exception("must not marshal Value2"); } set { throw new Exception("must not marshal Value2"); } }
    public object MergeCells
    {
        get
        {
            foreach (var m in Parent.Merges)
                if (m.Contains(Row,Column) || box.Contains(m.Row,m.Column)) return true;
            return false;
        }
    }
    public FakeRange MergeArea
    {
        get
        {
            foreach (var m in Parent.Merges) if (m.Contains(Row,Column)) return new FakeRange(Parent,m);
            return this;
        }
    }
    public static int CopyCalls;
    public void Copy() { CopyCalls++; clipboard=this; }
    public void UnMerge() { Parent.Merges.Clear(); }
    public void Merge() { if (!Parent.Merges.Contains(box)) Parent.Merges.Add(box); }
    public void PasteSpecial(int kind, object operation, object skipBlanks, object transpose)
    {
        if (kind == -4104 || kind == -4123) throw new Exception("all/formula paste forbidden");
        if (kind == 8) { Parent.WidthsCopied=true; return; }
        var from=clipboard;
        foreach (var merge in from.Parent.Merges)
        {
            if (!from.box.Contains(merge.Row,merge.Column)) continue;
            var corresponding=new Box(Row+merge.Row-from.Row,Column+merge.Column-from.Column,merge.Rows,merge.Columns);
            if (!Parent.Merges.Contains(corresponding)) throw new Exception("unequal merged-cell paste");
        }
        if (kind == -4122) return;
        if (kind != -4163) throw new Exception("unexpected paste kind");
        for (int r=0;r<from.box.Rows;r++) for (int c=0;c<from.box.Columns;c++) Parent.Values[Row+r,Column+c]=from.Parent.Values[from.Row+r,from.Column+c];
    }
}
