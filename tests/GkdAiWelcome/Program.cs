using System;
using System.Collections.Generic;
using System.Reflection;
using Outline.Addin;

public static class Program
{
    static int passed, failed;
    public static int Main()
    {
        try
        {
            Run("bare formula is rewritten with event restoration", () => {
                var app = new FakeApp(); var cell = new FakeCell("=GKD");
                Invoke(app, new FakeSheet(), new FakeRange(cell));
                Equal("=GKD()", cell.Formula); Equal(true, app.EnableEvents);
                Equal("False,True", string.Join(",", app.Changes));
            });
            Run("mixed ranges only rewrite exact case-insensitive trimmed formulas", () => {
                var app = new FakeApp();
                var cells = new[] { new FakeCell("  =gKd  "), new FakeCell("=GKD()"), new FakeCell("=GKD(\"hello\")"), new FakeCell(42), new FakeCell("=GKDX"), new FakeCell("=GKD+A1"), new FakeCell("=OtherName"), new FakeCell("'=GKD"), new FakeCell("= GKD"), new FakeCell("=GKD") { FormulaFlag = false } };
                Invoke(app, new FakeSheet(), new FakeRange(cells));
                Equal("=GKD()", cells[0].Formula);
                Equal(0, cells[9].Writes);
                for (int i = 1; i < cells.Length; i++) Equal(0, cells[i].Writes);
            });
            Run("range processes all matching cells", () => {
                var a = new FakeCell("=GKD"); var b = new FakeCell("=gkd");
                Invoke(new FakeApp(), new FakeSheet(), new FakeRange(a, b));
                Equal("=GKD()", a.Formula); Equal("=GKD()", b.Formula);
            });
            Run("protected worksheet is untouched", () => {
                var app = new FakeApp(); var cell = new FakeCell("=GKD");
                Invoke(app, new FakeSheet { ProtectContents = true }, new FakeRange(cell));
                Equal(0, cell.Writes); Equal(0, app.Changes.Count);
            });
            Run("invalid and oversized ranges are untouched", () => {
                foreach (double count in new[] { 0d, -1d, 10001d, double.NaN, double.PositiveInfinity }) {
                    var cell = new FakeCell("=GKD"); var app = new FakeApp();
                    Invoke(app, new FakeSheet(), new FakeRange(cell) { Size = count });
                    Equal(0, cell.Writes); Equal(0, app.Changes.Count);
                }
                Invoke(new FakeApp(), new FakeSheet(), null);
                Invoke(new FakeApp(), new FakeSheet(), new object());
            });
            Run("10000 cells accepted", () => {
                var cells = new FakeCell[10000]; for (int i = 0; i < cells.Length; i++) cells[i] = new FakeCell("=GKD");
                Invoke(new FakeApp(), new FakeSheet(), new FakeRange(cells));
                Equal(1, cells[9999].Writes);
            });
            Run("merged array and spill cells are untouched", () => {
                var cells = new[] { new FakeCell("=GKD") { MergeCells = true }, new FakeCell("=GKD") { HasArray = true }, new FakeCell("=GKD") { HasSpill = true } };
                var app = new FakeApp(); Invoke(app, new FakeSheet(), new FakeRange(cells));
                foreach (var c in cells) Equal(0, c.Writes);
                Equal(0, app.Changes.Count);
            });
            Run("event state false is preserved", () => {
                var app = new FakeApp(); app.EnableEvents = false; app.Changes.Clear();
                var cell = new FakeCell("=GKD"); Invoke(app, new FakeSheet(), new FakeRange(cell));
                Equal("=GKD()", cell.Formula); Equal(false, app.EnableEvents);
            });
            Run("write failure restores events and does not block other cells", () => {
                var a = new FakeCell("=GKD") { FailWrite = true }; var b = new FakeCell("=GKD"); var app = new FakeApp();
                Invoke(app, new FakeSheet(), new FakeRange(a, b));
                Equal(true, app.EnableEvents); Equal("=GKD()", b.Formula);
            });
            Run("subscription attaches once and stale callbacks are inert after stop", () => {
                var app = new FakeApp(); var events = new FakeEvents();
                Start(app, events); Start(app, events); Equal(1, events.Adds);
                var cell = new FakeCell("=GKD"); events.Fire(new FakeSheet(), new FakeRange(cell)); Equal(1, cell.Writes);
                Call("Stop"); Call("Stop"); Equal(1, events.Removes);
                var stale = new FakeCell("=GKD"); events.Fire(new FakeSheet(), new FakeRange(stale)); Equal(0, stale.Writes);
                Start(app, events); Equal(2, events.Adds); Call("Stop");
            });
            Run("failed subscribe and detach never escape and allow restart", () => {
                var app = new FakeApp(); var events = new FakeEvents { FailAdd = true };
                Start(app, events); var cell = new FakeCell("=GKD"); events.Fire(new FakeSheet(), new FakeRange(cell)); Equal(0, cell.Writes);
                events.FailAdd = false; Start(app, events); Equal(2, events.Adds);
                events.FailRemove = true; Call("Stop"); events.Fire(new FakeSheet(), new FakeRange(cell)); Equal(0, cell.Writes);
                events.FailRemove = false; Start(app, events); Call("Stop");
            });
            Run("EnableEvents disable failure prevents every write and restores original", () => {
                var app = new FakeApp { FailDisableOnce = true }; var a = new FakeCell("=GKD"); var b = new FakeCell("=GKD");
                Invoke(app, new FakeSheet(), new FakeRange(a, b));
                Equal(0, a.Writes); Equal(0, b.Writes); Equal(true, app.EnableEvents);
            });
            Run("existing local or workbook GKD name suppresses rewrite without enumerating names", () => {
                foreach (bool local in new[] { true, false }) {
                    var sheet = new FakeSheet(); if (local) sheet.Names.Exists = true; else sheet.Parent.Names.Exists = true;
                    var app = new FakeApp(); var cell = new FakeCell("=GKD"); Invoke(app, sheet, new FakeRange(cell));
                    Equal(0, cell.Writes); Equal(0, app.Changes.Count);
                }
            });
            Console.WriteLine("PASS " + passed + " tests; FAIL " + failed); return failed == 0 ? 0 : 1;
        }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
    }
    static void Invoke(FakeApp app, FakeSheet sheet, object range)
    {
        // Reflection allows RED to report the missing feature before it exists.
        Type type = typeof(ExcelDna.Integration.ExcelDnaUtil).Assembly.GetType("Outline.Addin.GkdAiWelcome");
        if (type == null) throw new Exception("Missing GkdAiWelcome interceptor");
        type.GetMethod("HandleChange", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new object[] { app, sheet, range });
    }
    static object Call(string method, params object[] args)
    {
        Type type = typeof(ExcelDna.Integration.ExcelDnaUtil).Assembly.GetType("Outline.Addin.GkdAiWelcome");
        MethodInfo found = Array.Find(type.GetMethods(BindingFlags.Static | BindingFlags.NonPublic), m => m.Name == method && m.GetParameters().Length == args.Length);
        if (found == null) throw new Exception("Missing " + method + " API");
        return found.Invoke(null, args);
    }
    static void Start(FakeApp app, FakeEvents events)
    {
        Call("Start", app, (Action<object, Guid, int, Delegate>)events.Add, (Action<object, Guid, int, Delegate>)events.Remove);
    }
    static void Run(string name, Action action) { try { action(); passed++; Console.WriteLine("PASS " + name); } catch (Exception e) { failed++; Console.WriteLine("FAIL " + name + ": " + e.Message); } }
    public static void Equal(object expected, object actual) { if (!Equals(expected, actual)) throw new Exception("Expected " + expected + ", got " + actual); }
}
public sealed class FakeEvents
{
    public int Adds, Removes;
    public bool FailAdd, FailRemove;
    readonly List<Delegate> callbacks = new List<Delegate>();
    public void Add(object app, Guid iid, int dispid, Delegate handler)
    {
        Program.Equal(new Guid("00024413-0000-0000-C000-000000000046"), iid); Program.Equal(0x61c, dispid);
        Adds++; callbacks.Add(handler); if (FailAdd) throw new InvalidOperationException("attach failed");
    }
    public void Remove(object app, Guid iid, int dispid, Delegate handler)
    {
        Removes++; if (FailRemove) throw new InvalidOperationException("detach failed");
        // Keep callbacks to simulate an already-dispatched callback after unload.
    }
    public void Fire(object sheet, object target) { foreach (var callback in callbacks.ToArray()) callback.DynamicInvoke(sheet, target); }
}
namespace Outline.Addin { }
public sealed class FakeApp
{
    bool enabled = true;
    public readonly List<bool> Changes = new List<bool>();
    public bool FailDisableOnce { get; set; }
    public bool EnableEvents { get { return enabled; } set { Changes.Add(value); enabled = value; if (!value && FailDisableOnce) { FailDisableOnce = false; throw new InvalidOperationException("disable failed"); } } }
}
public sealed class FakeNames
{
    public bool Exists { get; set; }
    public object Item(string name) { if (name.IndexOf("GKD", StringComparison.OrdinalIgnoreCase) < 0) throw new Exception("unexpected name lookup"); if (Exists) return new object(); throw new System.Runtime.InteropServices.COMException("name missing", unchecked((int)0x800A03EC)); }
}
public sealed class FakeSheet { public bool ProtectContents { get; set; } public FakeNames Names { get; } = new FakeNames(); public FakeWorkbook Parent { get; } = new FakeWorkbook(); }
public sealed class FakeWorkbook { public FakeNames Names { get; } = new FakeNames(); }
public sealed class FakeRange
{
    public FakeRange(params FakeCell[] cells) { Cells = new FakeCells(cells); Size = cells.Length; }
    public FakeCells Cells { get; private set; }
    public FakeAreas Areas { get { return new FakeAreas(this); } }
    public double Size { get; set; }
    public double CountLarge { get { return Size; } }
}
public sealed class FakeAreas
{
    readonly FakeRange range;
    public FakeAreas(FakeRange value) { range = value; }
    public int Count { get { return 1; } }
    public FakeRange this[int index] { get { return range; } }
}
public sealed class FakeCells
{
    readonly FakeCell[] cells;
    public FakeCells(FakeCell[] value) { cells = value; }
    public int Count { get { return cells.Length; } }
    public FakeCell this[int index] { get { return cells[index - 1]; } }
}
public sealed class FakeCell
{
    object formula;
    public FakeCell(object value) { formula = value; }
    public bool FailWrite { get; set; }
    public int Writes { get; private set; }
    public bool? FormulaFlag { get; set; }
    public object Formula { get { return formula; } set { if (FailWrite) throw new InvalidOperationException("write failed"); Writes++; formula = value; } }
    public bool HasFormula { get { return FormulaFlag ?? (Formula is string && ((string)Formula).TrimStart().StartsWith("=")); } }
    public bool MergeCells { get; set; }
    public bool HasArray { get; set; }
    public bool HasSpill { get; set; }
}
