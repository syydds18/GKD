using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Outline.Core;

// Runs the actual picker-reading methods against a portable Excel-shaped model.
// This exercises dynamic dispatch without pretending to run Excel COM or WinForms.
static class ReadChoicesTests
{
    internal static void Run(string source, string sourcePath)
    {
        string choice = Slice(source, "internal sealed class HydroSheetChoice", "internal sealed class HydropowerRangeForm");
        string methods = Slice(source, "internal static class HydropowerExtraction", "internal static string Run(");
        string harness = "using System; using System.Collections.Generic; using System.ComponentModel; using System.Globalization; using System.Linq; using Outline.Core; namespace Outline.Addin { " +
            choice + methods + " } }";
        string live = File.ReadAllText(Path.Combine(Path.GetDirectoryName(sourcePath), "HydroLivePrintArea.cs"));
        var references = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")).Split(Path.PathSeparator)
            .Concat(new[] { typeof(HydroExportPlanPolicy).Assembly.Location }).Distinct()
            .Select(x => MetadataReference.CreateFromFile(x));
        var compilation = CSharpCompilation.Create("PickerBehavior_" + Guid.NewGuid().ToString("N"),
            new[] { CSharpSyntaxTree.ParseText(harness), CSharpSyntaxTree.ParseText(live) }, references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        using (var stream = new MemoryStream())
        {
            var emitted = compilation.Emit(stream);
            if (!emitted.Success) throw new Exception(string.Join("\n", emitted.Diagnostics));
            var assembly = Assembly.Load(stream.ToArray());
            var read = assembly.GetType("Outline.Addin.HydropowerExtraction").GetMethod("ReadChoices", BindingFlags.NonPublic | BindingFlags.Static);
            var book = new PickerWorkbook();
            book.Worksheets.Add(new PickerSheet(book, "G1", "$B$2:$F$10", -1));
            book.Worksheets.Add(new PickerSheet(book, "封面", "$C$3:$E$9", -1));
            book.Worksheets.Add(new PickerSheet(book, "总概算", "$D$4:$G$8", -1));
            book.Worksheets.Add(new PickerSheet(book, "G6", "$A$1:$B$3", 0));
            var app = new PickerApp { ActiveWindow = new PickerWindow { SelectedSheets = new[] { book.Worksheets[2], book.Worksheets[1], book.Worksheets[3] } } };
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "picker-behavior-" + Guid.NewGuid().ToString("N") + ".xlsx");
            var store = HydroExportPlanStore.ForLocalApplicationData();
            string savedFile = (string)typeof(HydroExportPlanStore).GetMethod("GetPlanPath", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(store, new object[] { path });
            try
            {
                store.Save(path, new[] { new HydroExportPlanSheet { Name = "G1", PrintRange = "A1:A2" } });
                var result = ((IEnumerable)read.Invoke(null, new object[] { app, book, path })).Cast<object>().ToList();
                Check(result.Select(x => Text(x, "SheetName")).SequenceEqual(new[] { "G1", "封面", "总概算", "G6" }), "reading preserves workbook order regardless of Excel selection order");
                Check(result.Where(x => Flag(x, "Include")).Select(x => Text(x, "SheetName")).SequenceEqual(new[] { "封面", "总概算" }), "Excel multiselection wins; hidden sheets excluded");
                Check(Text(result[0], "PrintRange") == "B2:F10", "saved range never replaces current live range");
                book.Worksheets[0].PageSetup.PrintArea = "$E$5:$H$12";
                result = ((IEnumerable)read.Invoke(null, new object[] { app, book, path })).Cast<object>().ToList();
                Check(Text(result[0], "PrintRange") == "E5:H12", "reopening reads changed in-memory range");
                Console.WriteLine("PASS actual ReadChoices against portable Excel model (selection, live changes, hidden sheets, workbook order)");
            }
            finally
            {
                if (File.Exists(savedFile)) File.Delete(savedFile);
            }
        }
    }

    static string Slice(string source, string first, string last)
    {
        int start = source.IndexOf(first, StringComparison.Ordinal);
        return source.Substring(start, source.IndexOf(last, start, StringComparison.Ordinal) - start);
    }
    static bool Flag(object item, string name) { return (bool)item.GetType().GetProperty(name).GetValue(item); }
    static string Text(object item, string name) { return (string)item.GetType().GetProperty(name).GetValue(item); }
    static void Check(bool value, string message) { if (!value) throw new Exception("FAIL " + message); }
}

public sealed class PickerWorkbook { public List<PickerSheet> Worksheets { get; } = new List<PickerSheet>(); }
public sealed class PickerSheet
{
    public PickerSheet(PickerWorkbook parent, string name, string address, int visible)
    { Parent = parent; Name = name; Visible = visible; PageSetup = new PickerPageSetup { PrintArea = address }; }
    public PickerWorkbook Parent { get; }
    public string Name { get; }
    public int Visible { get; }
    public PickerPageSetup PageSetup { get; }
}
public sealed class PickerPageSetup { public string PrintArea { get; set; } }
public sealed class PickerWindow { public PickerSheet[] SelectedSheets { get; set; } }
public sealed class PickerApp
{
    public bool PrintCommunication { get; set; } = true;
    public PickerWindow ActiveWindow { get; set; }
}
