using System.Text.Json;
using System.Text.Json.Nodes;
using Outline.Core;

var options = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
var file = args.Length > 0 ? args[0] : Path.Combine(AppContext.BaseDirectory, "fixtures.json");
using var document = JsonDocument.Parse(File.ReadAllText(file));
var root = document.RootElement;
var sourceFile = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(file))!, "ExcelOutlineCore.original.js");
var sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(sourceFile))).ToLowerInvariant();
if (sourceHash != root.GetProperty("sourceSha256").GetString()) throw new InvalidOperationException("Original JS source SHA-256 differs from fixture provenance");
var failures = new List<string>();
int assertions = 0;
object? Value(JsonElement e) => e.ValueKind switch { JsonValueKind.Number => e.GetDouble(), JsonValueKind.String => e.GetString(), JsonValueKind.True => true, JsonValueKind.False => false, _ => null };
void Compare(JsonElement expected, object actual, string name)
{
    var observed = JsonSerializer.SerializeToElement(actual, options);
    Walk(expected, observed, name);
}
void Walk(JsonElement expected, JsonElement actual, string name)
{
    assertions++;
    // The original JS omits optional fields. Every field it actually emitted must match.
    if (expected.ValueKind == JsonValueKind.Object && actual.ValueKind == JsonValueKind.Object)
    {
        foreach(var p in expected.EnumerateObject())
        {
            if(!actual.TryGetProperty(p.Name,out var a)) failures.Add(name+"."+p.Name+": missing");
            else Walk(p.Value,a,name+"."+p.Name);
        }
    }
    else if(expected.ValueKind == JsonValueKind.Array && actual.ValueKind == JsonValueKind.Array)
    {
        if(expected.GetArrayLength()!=actual.GetArrayLength()) {failures.Add(name+": array lengths differ "+expected.GetArrayLength()+" / "+actual.GetArrayLength());return;}
        for(int i=0;i<expected.GetArrayLength();i++) Walk(expected[i],actual[i],name+"["+i+"]");
    }
    else if(expected.ValueKind==JsonValueKind.Number && actual.ValueKind==JsonValueKind.Number)
    {
        double a=expected.GetDouble(), b=actual.GetDouble();
        if(a!=b) failures.Add(name+": "+a+" != "+b);
    }
    else if(expected.ToString()!=actual.ToString() || expected.ValueKind!=actual.ValueKind) failures.Add(name+": "+expected.GetRawText()+" != "+actual.GetRawText());
}
foreach(var test in root.GetProperty("codeCases").EnumerateArray()) Compare(test.GetProperty("expected"),Engine.CodeFamily(Value(test.GetProperty("input"))),"code:"+test.GetProperty("input"));
foreach(var test in root.GetProperty("numberCases").EnumerateArray()) Compare(test.GetProperty("expected"),Engine.ToNumberValue(Value(test.GetProperty("input"))),"number:"+test.GetProperty("input"));
foreach(var test in root.GetProperty("sheets").EnumerateArray())
{
    var input=test.GetProperty("input");
    var sheet=new SheetData {MaxRow=input.GetProperty("maxRow").GetInt32(),MaxCol=input.GetProperty("maxCol").GetInt32(),OutlineMax=input.GetProperty("outlineMax").GetInt32()};
    foreach(var r in input.GetProperty("values").EnumerateObject())
    {
        var row=new Dictionary<int,object>();
        foreach(var c in r.Value.EnumerateObject()) row[int.Parse(c.Name)]=Value(c.Value)!;
        sheet.Values[int.Parse(r.Name)]=row;
    }
    foreach(var r in input.GetProperty("formulas").EnumerateObject())
    {
        var row=new Dictionary<int,string>();
        foreach(var c in r.Value.EnumerateObject()) row[int.Parse(c.Name)]=c.Value.GetString()!;
        sheet.Formulas[int.Parse(r.Name)]=row;
    }
    foreach(var r in input.GetProperty("outlineLevels").EnumerateObject()) sheet.OutlineLevels[int.Parse(r.Name)]=r.Value.GetInt32();
    var detection=Engine.Detect(sheet);
    var name=test.GetProperty("name").GetString()!;
    Compare(test.GetProperty("header"),Engine.FindHeader(sheet),name+".header");
    Compare(test.GetProperty("rawDetection"),Engine.DetectLevels(sheet),name+".rawDetection");
    Compare(test.GetProperty("detection"),detection,name+".detection");
    Compare(test.GetProperty("validation"),Engine.Validate(sheet,detection),name+".validation");
}
if(failures.Count>0) {foreach(var failure in failures.Take(40))Console.Error.WriteLine(failure);Console.Error.WriteLine($"FAIL: {failures.Count} mismatches, {assertions} assertions");return 1;}
Console.WriteLine($"PASS: {root.GetProperty("sheets").GetArrayLength()} sheets, {root.GetProperty("codeCases").GetArrayLength()} code cases, {root.GetProperty("numberCases").GetArrayLength()} number cases, {assertions} assertions; source SHA-256 {root.GetProperty("sourceSha256").GetString()}");
return 0;
