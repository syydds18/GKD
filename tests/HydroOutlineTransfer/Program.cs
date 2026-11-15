using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

static class Program
{
    static int Main()
    {
        try { Run(); return 0; }
        catch (Exception ex) { Console.Error.WriteLine("FAIL: " + ex); return 1; }
    }
    static void Copy(FakeRange source, FakeSheet destination)
    {
        var type = typeof(Program).Assembly.GetType("Outline.Addin.HydroOutlineTransfer");
        Check(type != null, "HydroOutlineTransfer helper is missing");
        var method = type.GetMethod("Copy", BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static);
        Check(method != null, "Copy(sourceRange, destinationSheet) is missing");
        try { method.Invoke(null, new object[] { source, destination }); }
        catch (TargetInvocationException ex) { throw ex.InnerException; }
    }
    static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
    static void Run()
    {
        var source = new FakeSheet(20, 12);
        source.RowLevels = new[] {1,1,2,2,3,3,2,1,2,2,2,1,1,1,1,1,1,1,1,1};
        source.ColLevels = new[] {1,1,2,3,3,2,1,2,2,1,1,1};
        source.RowHidden[4] = source.RowHidden[5] = true;
        source.ColHidden[3] = source.ColHidden[4] = true;
        source.RowHidden[9] = true; // manual hiding in a different group
        source.Outline.SummaryRow = 0;
        source.Outline.SummaryColumn = -4131;
        var expectedRows = source.RowLevels.Skip(3).Take(8).ToArray();
        var expectedCols = source.ColLevels.Skip(2).Take(7).ToArray();
        var rowsBefore = source.RowLevels.ToArray();
        var colsBefore = source.ColLevels.ToArray();
        var destination = new FakeSheet(20, 12);
        Copy(new FakeRange(source, 4, 3, 8, 7), destination);
        Check(destination.RowLevels.Take(8).SequenceEqual(expectedRows), "clipped row outline levels changed");
        Check(destination.ColLevels.Take(7).SequenceEqual(expectedCols), "clipped column outline levels changed");
        Check(destination.RowHidden.Take(8).SequenceEqual(source.RowHidden.Skip(3).Take(8)), "row hidden states not mapped to A1");
        Check(destination.ColHidden.Take(7).SequenceEqual(source.ColHidden.Skip(2).Take(7)), "column hidden states not mapped to A1");
        Check(destination.Outline.SummaryRow == 0 && destination.Outline.SummaryColumn == -4131, "summary directions changed");
        Check(source.Writes == 0 && source.RowLevels.SequenceEqual(rowsBefore) && source.ColLevels.SequenceEqual(colsBefore), "source was modified");
        Check(destination.RowLevels.Skip(8).All(x=>x==1) && destination.ColLevels.Skip(7).All(x=>x==1), "outline leaked outside clipped destination");
        Check(destination.Groups.SequenceEqual(new[] {"R:1:4", "R:6:8", "R:2:3", "C:1:4", "C:6:7", "C:2:3"}), "groups must be outer-to-inner in both axes");
        Console.WriteLine("PASS clipped nested row/column levels, hidden states, summary directions, source immutability");
        TestValidationAndBatching();
        TestCollapsedSummaries();
    }
    static void TestCollapsedSummaries()
    {
        foreach (bool before in new[] {false,true})
        {
            var source=new FakeSheet(12,12); var target=new FakeSheet(12,12);
            source.Outline.SummaryRow=before?0:1;
            source.Outline.SummaryColumn=before?-4131:-4152;
            source.RowLevels=new[] {1,2,2,3,3,2,1,1,1,1,1,1};
            source.ColLevels=source.RowLevels.ToArray();
            int outer=before?1:7, inner=before?3:6;
            source.Details["R:"+outer]=false; source.Details["R:"+inner]=true;
            source.Details["C:"+outer]=true; source.Details["C:"+inner]=false;
            source.RowHidden[1]=true; source.RowHidden[3]=true; source.ColHidden[4]=true;
            Copy(new FakeRange(source,1,1,8,8),target);
            Check(target.Details.Count==4,"summary ShowDetail states not transferred");
            foreach(var pair in source.Details) Check(target.Details.ContainsKey(pair.Key) && target.Details[pair.Key]==pair.Value,"summary collapsed state changed: "+pair.Key);
            Check(target.RowHidden.Take(8).SequenceEqual(source.RowHidden.Take(8)) && target.ColHidden.Take(8).SequenceEqual(source.ColHidden.Take(8)),"ShowDetail must not overwrite final manual hidden states");
            Check(source.Writes==0,"summary transfer mutated source");
        }
        var clipped=new FakeSheet(8,8); var clippedTarget=new FakeSheet(8,8);
        clipped.RowLevels=Enumerable.Repeat(3,8).ToArray(); clipped.ColLevels=Enumerable.Repeat(2,8).ToArray();
        Copy(new FakeRange(clipped,2,2,5,5),clippedTarget);
        Check(clipped.DetailReads==0,"clipped-out summaries must not be read or fabricated");
        var unavailable=new FakeSheet(8,8); unavailable.RowLevels[1]=2;
        Copy(new FakeRange(unavailable,1,1,8,8),new FakeSheet(8,8));
        var failedTarget=new FakeSheet(8,8) { RejectDetails=true };
        unavailable.Details["R:3"]=false; unavailable.RowHidden[1]=true;
        Copy(new FakeRange(unavailable,1,1,8,8),failedTarget);
        Check(failedTarget.RowHidden[1],"failed best-effort ShowDetail must still restore hidden states");
        Console.WriteLine("PASS above/below and left/right ShowDetail, unavailable/clipped summaries, hidden restored last");
    }
    static void Reject(FakeRange source, FakeSheet destination, string name)
    {
        int writes=destination.Writes;
        try { Copy(source,destination); }
        catch (ArgumentException) { Check(destination.Writes==writes,name+" wrote destination before validation"); return; }
        catch (InvalidOperationException) { Check(destination.Writes==writes,name+" wrote destination before validation"); return; }
        throw new Exception(name+" must be rejected");
    }
    static void TestValidationAndBatching()
    {
        var source=new FakeSheet(30,20); var target=new FakeSheet(30,20);
        Reject(new FakeRange(source,1,1,0,1),target,"empty row extent");
        Reject(new FakeRange(source,1,1,1,0),target,"empty column extent");
        Reject(new FakeRange(source,1,1,1048577,1),target,"row count limit");
        Reject(new FakeRange(source,1,1,1,16385),target,"column count limit");
        Reject(new FakeRange(source,1048576,1,2,1),target,"source row boundary");
        Reject(new FakeRange(source,1,16384,1,2),target,"source column boundary");
        Reject(new FakeRange(source,0,1,1,1),target,"invalid source origin");
        Reject(new FakeRange(source,1,1,4,4) { Areas=new FakeCount(2) },target,"multi area range");
        Reject(new FakeRange(source,1,1,4,4),source,"same worksheet");
        source.ColLevels[3]=9;
        Reject(new FakeRange(source,1,1,4,4),target,"invalid source depth");
        source.ColLevels[3]=1; source.RowLevels[1]=0;
        Reject(new FakeRange(source,1,1,4,4),target,"invalid source zero depth");
        source.RowLevels[1]=1; source.Outline.SummaryRow=2;
        Reject(new FakeRange(source,1,1,4,4),target,"invalid summary direction");
        source.Outline.SummaryRow=1; target.RowLevels[1]=2;
        Reject(new FakeRange(source,1,1,4,4),target,"destination must be fresh");
        source=new FakeSheet(1048576,16384); target=new FakeSheet(1048576,16384);
        Copy(new FakeRange(source,1,1,1048576,16384),target);
        Check(source.LevelReads==2 && source.HiddenReads==2,"uniform axes must use one range read per property, not per row/cell");
        Check(target.Groups.Count==0 && target.Writes==2,"uniform hidden states must use one assignment per axis");
        source=new FakeSheet(16,16); target=new FakeSheet(16,16);
        for(int i=0;i<16;i++) { source.RowLevels[i]=i%8+1; source.ColLevels[i]=8-i%8; source.RowHidden[i]=i%2==0; source.ColHidden[i]=i%3==0; }
        Copy(new FakeRange(source,1,1,16,16),target);
        Check(source.RowLevels.SequenceEqual(target.RowLevels) && source.ColLevels.SequenceEqual(target.ColLevels),"all eight levels must reconstruct");
        Check(source.RowHidden.SequenceEqual(target.RowHidden) && source.ColHidden.SequenceEqual(target.ColHidden),"varied singleton hidden fallback changed values");
        Check(source.LevelReads<=62 && source.HiddenReads<=62,"recursive fallback read budget exceeded");
        Console.WriteLine("PASS extent/depth/fresh-target validation, full-sheet uniform batching, varied eight-level fallback");
    }
}

public sealed class FakeSheet
{
    public int[] RowLevels, ColLevels;
    public bool[] RowHidden, ColHidden;
    public readonly FakeOutline Outline = new FakeOutline();
    public int Writes, LevelReads, HiddenReads, DetailReads;
    public bool RejectDetails;
    public readonly Dictionary<string,bool> Details = new Dictionary<string,bool>();
    public readonly List<string> Groups = new List<string>();
    public FakeSheet(int rows, int columns)
    {
        RowLevels=Enumerable.Repeat(1,rows).ToArray(); ColLevels=Enumerable.Repeat(1,columns).ToArray();
        RowHidden=new bool[rows]; ColHidden=new bool[columns];
    }
    public FakeCells Cells { get { return new FakeCells(this); } }
    public FakeRangeIndexer Range { get { return new FakeRangeIndexer(this); } }
    public FakeRows Rows { get { return new FakeRows(this); } }
}
public sealed class FakeOutline { public int SummaryRow { get; set; } = 1; public int SummaryColumn { get; set; } = -4152; }
public sealed class FakeCount { public int Count { get; } public FakeCount(int count) { Count=count; } }
public sealed class FakeCells
{
    readonly FakeSheet sheet;
    public FakeCells(FakeSheet s) { sheet=s; }
    public FakeRange this[int row,int column] { get { return new FakeRange(sheet,row,column,1,1); } }
}
public sealed class FakeRangeIndexer
{
    readonly FakeSheet sheet;
    public FakeRangeIndexer(FakeSheet s) { sheet=s; }
    public FakeRange this[FakeRange first,FakeRange last] { get { return new FakeRange(sheet,first.Row,first.Column,last.Row-first.Row+1,last.Column-first.Column+1); } }
}
public sealed class FakeRows
{
    readonly FakeSheet sheet;
    public FakeRows(FakeSheet s) { sheet=s; }
    public FakeAxis this[string address]
    {
        get { var parts=address.Split(':'); return new FakeAxis(sheet,false,int.Parse(parts[0]),int.Parse(parts[1])); }
    }
}
public sealed class FakeRange
{
    public FakeSheet Parent { get; }
    public int Row { get; } public int Column { get; }
    public FakeCount Rows { get; } public FakeCount Columns { get; }
    public FakeCount Areas { get; set; } = new FakeCount(1);
    public FakeRange(FakeSheet sheet,int row,int column,int rows,int columns) { Parent=sheet; Row=row; Column=column; Rows=new FakeCount(rows); Columns=new FakeCount(columns); }
    public FakeAxis EntireRow { get { return new FakeAxis(Parent,false,Row,Row+Rows.Count-1); } }
    public FakeAxis EntireColumn { get { return new FakeAxis(Parent,true,Column,Column+Columns.Count-1); } }
}
public sealed class FakeAxis
{
    readonly FakeSheet sheet; readonly bool columns; readonly int first,last;
    int[] Levels { get { return columns?sheet.ColLevels:sheet.RowLevels; } }
    bool[] HiddenStates { get { return columns?sheet.ColHidden:sheet.RowHidden; } }
    public FakeAxis(FakeSheet s,bool c,int f,int l) { sheet=s; columns=c; first=f; last=l; }
    public object OutlineLevel
    {
        get { sheet.LevelReads++; int value=Levels[first-1]; for(int i=first;i<last;i++) if(Levels[i]!=value) return DBNull.Value; return (double)value; }
    }
    public object Hidden
    {
        get { sheet.HiddenReads++; bool value=HiddenStates[first-1]; for(int i=first;i<last;i++) if(HiddenStates[i]!=value) return null; return value; }
        set { sheet.Writes++; for(int i=first-1;i<last;i++) HiddenStates[i]=(bool)value; }
    }
    public object ShowDetail
    {
        get
        {
            sheet.DetailReads++; bool state;
            if(first!=last || !sheet.Details.TryGetValue((columns?"C:":"R:")+first,out state)) throw new System.Runtime.InteropServices.COMException("not a summary");
            return state;
        }
        set
        {
            if(sheet.RejectDetails) throw new System.Runtime.InteropServices.COMException("summary not available");
            sheet.Writes++;
            sheet.Details[(columns?"C:":"R:")+first]=(bool)value;
            // Model Excel ShowDetail side effects to prove final Hidden restoration.
            for(int i=0;i<HiddenStates.Length;i++) HiddenStates[i]=!(bool)value;
        }
    }
    public void Group()
    {
        sheet.Writes++; sheet.Groups.Add((columns?"C:":"R:")+first+":"+last);
        for(int i=first-1;i<last;i++) { if(Levels[i]>=8) throw new Exception("Excel outline depth exceeded"); Levels[i]++; }
    }
}
