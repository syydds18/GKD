using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using Newtonsoft.Json.Linq;
using Outline.Addin;

internal static class Program
{
    private static int passed, failed;
    private static Type FormulaType { get { return typeof(Program).Assembly.GetType("Outline.Addin.GkdCustomFormula"); } }
    private static object Formula()
    {
        if (FormulaType == null) throw new Exception("custom formula model missing");
        var f = Activator.CreateInstance(FormulaType);
        Set(f, "ExcelName", "GKD_AREA"); Set(f, "DisplayName", "面积");
        Set(f, "Description", "计算矩形面积"); Set(f, "FormulaExpression", "2 * 3");
        return f;
    }
    private static void Set(object value, string name, object data) { value.GetType().GetProperty(name).SetValue(value, data); }
    private static object Get(object value, string name) { return value.GetType().GetProperty(name).GetValue(value); }
    private static object Call(object value, string method, params object[] args)
    {
        try { return value.GetType().GetMethod(method).Invoke(value, args); }
        catch (TargetInvocationException e) { throw e.InnerException; }
    }
    private static void Equal(object expected, object actual) { if (!Equals(expected, actual)) throw new Exception("expected " + expected + ", got " + actual); }
    private static void Reject(Action action)
    {
        try { action(); } catch (ArgumentException) { return; } catch (InvalidDataException) { return; }
        throw new Exception("expected validation rejection");
    }
    private static void Run(string name, Action test)
    {
        try { test(); passed++; Console.WriteLine("PASS " + name); }
        catch (Exception e) { failed++; Console.WriteLine("FAIL " + name + ": " + e.Message); }
    }
    private static void Main()
    {
        Run("numeric grammar, precedence and explicit input binding", Arithmetic);
        Run("unsafe syntax, undeclared inputs and bounded parsing", ExpressionSafety);
        Run("allowlisted dimension-compatible unit conversion", Units);
        Run("definition defaults and strict metadata validation", Metadata);
        Run("input declarations and kind-specific requirements", DefinitionKinds);
        Run("versioned JSON export/import preserves data without credentials", JsonRoundTrip);
        Run("strict JSON schema, types, duplicate identity and size limits", JsonStrict);
        Run("literal validation and compatible definition boundaries", Boundaries);
        Console.WriteLine("RESULT: " + passed + " passed, " + failed + " failed");
        if (failed != 0) Environment.ExitCode = 1;
    }
    private static double Evaluate(string expression, IDictionary<string,double> inputs = null)
    {
        var f = Formula(); Set(f,"FormulaExpression",expression);
        Set(f,"Inputs", inputs == null ? new List<string>() : new List<string>(inputs.Keys));
        return (double)Call(f,"Evaluate",inputs ?? new Dictionary<string,double>());
    }
    private static void Arithmetic()
    {
        if (FormulaType == null || FormulaType.GetMethod("Evaluate") == null) throw new Exception("safe evaluator missing");
        Equal(14.0, Evaluate("2 + 3 * 4")); Equal(20.0, Evaluate("(2+3)*4"));
        Equal(512.0, Evaluate("2^3^2")); Equal(-4.0, Evaluate("-2^2")); Equal(0.25,Evaluate("2^-2"));
        Equal(3.0, Evaluate("1e1 / 4 + .5")); Equal(2.0,Evaluate("--2"));
        var inputs = new Dictionary<string,double>{{"WIDTH",2.5},{"HEIGHT",4}};
        Equal(10.0,Evaluate("WIDTH * HEIGHT",inputs));
        var f = Formula(); Set(f,"Inputs",new List<string>{"X"}); Set(f,"FormulaExpression","X+1");
        Reject(() => Call(f,"Evaluate",new Dictionary<string,double>()));
        Reject(() => Call(f,"Evaluate",new Dictionary<string,double>{{"X",1},{"Y",2}}));
        Reject(() => Call(f,"Evaluate",new Dictionary<string,double>{{"X",double.NaN}}));
        Reject(() => Call(f,"Evaluate",new Dictionary<string,double>{{"X",double.PositiveInfinity}}));
        Reject(() => Call(f,"Evaluate",new object[]{null}));
        Set(f,"Enabled",false); Reject(() => Call(f,"Evaluate",new Dictionary<string,double>{{"X",1}}));
    }
    private static void ExpressionSafety()
    {
        foreach (string expression in new[] { "UNDECLARED+1", "=1+1", "1;2", "Math.Pow(2,3)", "System.IO.File.Delete(\"x\")", "SUM(1,2)", "1/*comment*/+2", "[A1]", "A1:B3", "1+", "(1+2", "1 2", "1e", "0x10", "NaN", "Infinity", "1==1", "\"hello\"" })
        {
            var f = Formula(); Set(f,"FormulaExpression",expression); Reject(() => Call(f,"Validate"));
            Reject(() => Evaluate(expression));
        }
        foreach (string expression in new[] { new string('(',65)+"1"+new string(')',65), new string('-',65)+"1", "2^"+new string('-',65)+"2" })
        { var f = Formula(); Set(f,"FormulaExpression",expression); Reject(() => Call(f,"Validate")); }
        foreach (string expression in new[] { "1/0", "0/0", "1e309", "1e308*1e308", "(-1)^.5", "2^1024" }) Reject(() => Evaluate(expression));
        var symbolic = Formula(); Set(symbolic,"Inputs",new List<string>{"X"}); Set(symbolic,"FormulaExpression","1 / (X-X)"); Call(symbolic,"Validate");
        Reject(() => Call(symbolic,"Evaluate",new Dictionary<string,double>{{"X",4}}));
        var ignoreCase = new Dictionary<string,double>(StringComparer.OrdinalIgnoreCase){{"x",1}};
        Set(symbolic,"FormulaExpression","X"); Reject(() => Call(symbolic,"Evaluate",ignoreCase));
        var old = System.Globalization.CultureInfo.CurrentCulture;
        try { System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("fr-FR"); Equal(1.75,Evaluate("1.25+.5")); Reject(() => Evaluate("1,25")); }
        finally { System.Globalization.CultureInfo.CurrentCulture = old; }
    }
    private static void Units()
    {
        Equal(1.5, Evaluate("CONVERT(1500,\"mm\",\"m\")"));
        Equal(1000.0, Evaluate("CONVERT(1,\"t\",\"kg\")"));
        Equal(2000000.0, Evaluate("CONVERT(2,\"MPa\",\"Pa\")"));
        Equal(120.0, Evaluate("CONVERT(2,\"h\",\"min\")"));
        Equal(2.0, Evaluate("CONVERT(20000,\"cm2\",\"m2\")"));
        Equal(1.0, Evaluate("CONVERT(1000,\"L\",\"m3\")"));
        Equal(2.0, Evaluate("CONVERT(CONVERT(2,\"m\",\"mm\"),\"mm\",\"m\")"));
        foreach (string expression in new[] { "CONVERT(1,\"m\",\"kg\")", "CONVERT(1,\"unknown\",\"m\")", "CONVERT(1,\"M\",\"m\")", "CONVERT(1,\"m\")", "CONVERT(1,'m','mm')", "CONVERT(1,X,\"m\")" })
        { var f = Formula(); Set(f,"FormulaExpression",expression); Reject(() => Call(f,"Validate")); }
        Reject(() => Evaluate("CONVERT(1e308,\"km\",\"mm\")"));
    }
    private static void DefinitionKinds()
    {
        var f = Formula(); Set(f, "Inputs", null); Reject(() => Call(f, "Validate"));
        foreach (var inputs in new[] { new List<string>{"X","X"}, new List<string>{"x"}, new List<string>{""}, new List<string>{null}, new List<string>{"A B"} })
        { f = Formula(); Set(f, "Inputs", inputs); Reject(() => Call(f, "Validate")); }
        var many = new List<string>(); for (int i = 0; i < 33; i++) many.Add("X" + i);
        f = Formula(); Set(f, "Inputs", many); Reject(() => Call(f, "Validate"));
        f = Formula(); Set(f, "FormulaExpression", " "); Reject(() => Call(f, "Validate"));
        f = Formula(); Set(f, "FormulaExpression", new string('1',4097)); Reject(() => Call(f, "Validate"));
        f = Formula(); Set(f, "PromptTemplate", "unexpected prompt"); Reject(() => Call(f, "Validate"));
        f = Formula(); Set(f, "Kind", Enum.Parse(Get(f,"Kind").GetType(), "Ai")); Set(f,"FormulaExpression", "");
        Reject(() => Call(f,"Validate"));
        Set(f, "PromptTemplate", "分析 {DATA}\n给出解释"); Call(f,"Validate");
        Set(f,"FormulaExpression", "2+3"); Reject(() => Call(f,"Validate")); Set(f,"FormulaExpression", "");
        Set(f,"PromptTemplate", new string('x',16001)); Reject(() => Call(f,"Validate"));
        Set(f,"PromptTemplate", "bad\0prompt"); Reject(() => Call(f,"Validate"));
    }
    private static object JsonCall(string method, object value)
    {
        var type = typeof(Program).Assembly.GetType("Outline.Addin.GkdCustomFormulaJson");
        if (type == null) throw new Exception("JSON contract missing");
        try { return type.GetMethod(method).Invoke(null,new[]{value}); }
        catch (TargetInvocationException e) { throw e.InnerException; }
    }
    private static object Definitions(params object[] formulas)
    {
        var list = (System.Collections.IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(FormulaType));
        foreach (object f in formulas) list.Add(f); return list;
    }
    private static void JsonRoundTrip()
    {
        var f = Formula(); Set(f,"Inputs",new List<string>{"X"}); Set(f,"FormulaExpression","X*2"); Set(f,"OutputUnit","m2"); Set(f,"Version",3); Set(f,"Enabled",false);
        var ai = Formula(); Set(ai,"ExcelName","GKD_EXPLAIN"); Set(ai,"Kind",Enum.Parse(Get(ai,"Kind").GetType(),"Ai")); Set(ai,"FormulaExpression",""); Set(ai,"PromptTemplate","解释数据\n保留原意");
        string json = (string)JsonCall("Export",Definitions(f,ai)); var root = JObject.Parse(json);
        Equal(1,(int)root["SchemaVersion"]); Equal(2,((JArray)root["Formulas"]).Count);
        Equal("Deterministic",(string)root["Formulas"][0]["Kind"]); Equal("Ai",(string)root["Formulas"][1]["Kind"]);
        Equal(11,((JObject)root["Formulas"][0]).Count);
        foreach (string forbidden in new[]{"ApiKey","Endpoint","Provider","$type"}) Equal(false,json.Contains(forbidden));
        var imported = (System.Collections.IList)JsonCall("Import",json);
        Equal(2,imported.Count); Equal(Get(f,"Id"),Get(imported[0],"Id")); Equal(3,Get(imported[0],"Version")); Equal(false,Get(imported[0],"Enabled"));
        Equal("m2",Get(imported[0],"OutputUnit")); Equal("X",((List<string>)Get(imported[0],"Inputs"))[0]); Equal("解释数据\n保留原意",Get(imported[1],"PromptTemplate"));
        Equal(json,JsonCall("Export",imported));
        ((List<string>)Get(imported[0],"Inputs")).Add("Y"); Equal(1,((List<string>)Get(f,"Inputs")).Count);
        Equal(0,((System.Collections.IList)JsonCall("Import",JsonCall("Export",Definitions()))).Count);
    }
    private static void JsonStrict()
    {
        string json = (string)JsonCall("Export",Definitions(Formula()));
        var root = JObject.Parse(json);
        foreach (string field in new[]{"Id","ExcelName","DisplayName","Description","Kind","Inputs","FormulaExpression","PromptTemplate","OutputUnit","Version","Enabled"})
        {
            var changed = (JObject)root.DeepClone(); ((JObject)changed["Formulas"][0]).Remove(field); Reject(() => JsonCall("Import",changed.ToString()));
            changed = (JObject)root.DeepClone(); changed["Formulas"][0][field] = JValue.CreateNull(); Reject(() => JsonCall("Import",changed.ToString()));
            changed = (JObject)root.DeepClone(); changed["Formulas"][0][field] = new JObject(); Reject(() => JsonCall("Import",changed.ToString()));
        }
        foreach (string field in new[]{"Id","ExcelName","Description","Kind","FormulaExpression","PromptTemplate","OutputUnit"})
        { var c = (JObject)root.DeepClone(); c["Formulas"][0][field] = 1; Reject(() => JsonCall("Import",c.ToString())); }
        foreach (string field in new[]{"SchemaVersion","Version"})
        foreach (JToken bad in new JToken[]{new JValue("1"),new JValue(1.0),new JValue(true),new JValue(2147483648L)})
        { var c = (JObject)root.DeepClone(); if(field=="SchemaVersion") c[field]=bad; else c["Formulas"][0][field]=bad; Reject(() => JsonCall("Import",c.ToString())); }
        foreach (JToken bad in new JToken[]{new JValue(1),new JValue("true")})
        { var c = (JObject)root.DeepClone(); c["Formulas"][0]["Enabled"]=bad; Reject(() => JsonCall("Import",c.ToString())); }
        foreach (string key in new[]{"ApiKey","$type","Unknown"})
        {
            var c=(JObject)root.DeepClone(); c[key]="injected"; Reject(() => JsonCall("Import",c.ToString()));
            c=(JObject)root.DeepClone(); c["Formulas"][0][key]="injected"; Reject(() => JsonCall("Import",c.ToString()));
        }
        foreach (string bad in new[]{"", "null", "[]", "{}", "{\"SchemaVersion\":2,\"Formulas\":[]}", "{\"SchemaVersion\":1}", "{\"SchemaVersion\":1,\"Formulas\":[null]}", "{\"SchemaVersion\":1,\"Formulas\":[1]}", "{\"SchemaVersion\":1,\"Formulas\":null}", "{\"SchemaVersion\":1,\"Formulas\":[],}", "{SchemaVersion:1,Formulas:[]}", "{\"SchemaVersion\":1,/*c*/\"Formulas\":[]}", "{\"SchemaVersion\":1,\"SchemaVersion\":1,\"Formulas\":[]}"}) Reject(() => JsonCall("Import",bad));
        Reject(() => JsonCall("Import",json+" trailing")); Reject(() => JsonCall("Import",new string(' ',1048577)));
        var duplicate=(JObject)root.DeepClone(); ((JArray)duplicate["Formulas"]).Add(duplicate["Formulas"][0].DeepClone()); Reject(() => JsonCall("Import",duplicate.ToString()));
        var first=Formula(); var second=Formula(); Reject(() => JsonCall("Export",Definitions(first,second)));
        Set(second,"ExcelName","GKD_OTHER"); Set(second,"Id",Get(first,"Id")); Reject(() => JsonCall("Export",Definitions(first,second)));
        var upper=Formula(); Set(upper,"Id",((string)Get(first,"Id")).ToUpperInvariant()); Set(upper,"ExcelName","GKD_UPPER"); Reject(() => JsonCall("Export",Definitions(first,upper)));
        Reject(() => JsonCall("Export",null)); Reject(() => JsonCall("Export",Definitions(new object[]{null})));
        var list=(System.Collections.IList)Definitions();
        for(int i=0;i<257;i++){var f=Formula();Set(f,"ExcelName","F"+i);list.Add(f);} Reject(() => JsonCall("Export",list));
        var huge = (JObject)root.DeepClone(); var arr=(JArray)huge["Formulas"]; arr.Clear();
        for(int i=0;i<257;i++){var f=(JObject)root["Formulas"][0].DeepClone();f["Id"]=Guid.NewGuid().ToString("D");f["ExcelName"]="F"+i;arr.Add(f);} Reject(() => JsonCall("Import",huge.ToString()));
    }
    private static void Boundaries()
    {
        var f=Formula(); Set(f,"Id"," "+Get(f,"Id")+" "); Reject(() => Call(f,"Validate"));
        f=Formula(); Set(f,"FormulaExpression","1e309"); Reject(() => Call(f,"Validate"));
        f=Formula(); Set(f,"FormulaExpression",new string('9',4000)); Reject(() => Call(f,"Validate"));
        f=Formula(); Set(f,"ExcelName","A"+new string('Z',30)); Set(f,"DisplayName",new string('x',100)); Set(f,"Description",new string('x',4000)); Set(f,"OutputUnit",new string('x',32)); Call(f,"Validate");
        f=Formula(); Set(f,"Kind",Enum.Parse(Get(f,"Kind").GetType(),"Ai")); Set(f,"FormulaExpression",""); Set(f,"PromptTemplate","test"); Reject(() => Call(f,"Evaluate",new Dictionary<string,double>()));
        var list=(System.Collections.IList)Definitions();
        for(int i=0;i<256;i++){var item=Formula();Set(item,"ExcelName","F"+i);list.Add(item);}
        string boundary=(string)JsonCall("Export",list); Equal(256,((System.Collections.IList)JsonCall("Import",boundary)).Count);
        foreach(object item in list) Set(item,"Description",new string('x',4000)); Reject(() => JsonCall("Export",list));
        string json=(string)JsonCall("Export",Definitions(Formula())); var root=JObject.Parse(json);
        foreach(JToken bad in new JToken[]{new JValue(1),JValue.CreateNull(),new JObject()})
        { var c=(JObject)root.DeepClone(); c["Formulas"][0]["Inputs"]=new JArray(bad); Reject(() => JsonCall("Import",c.ToString())); }
        foreach(string kind in new[]{"ai","1","0","Script"})
        { var c=(JObject)root.DeepClone(); c["Formulas"][0]["Kind"]=kind; Reject(() => JsonCall("Import",c.ToString())); }
        Reject(() => JsonCall("Import",null));
    }
    private static void Metadata()
    {
        var f = Formula(); Call(f, "Validate");
        Guid id; Equal(true, Guid.TryParseExact((string)Get(f, "Id"), "D", out id) && id != Guid.Empty);
        Equal(1, Get(f, "Version")); Equal(true, Get(f, "Enabled")); Equal("Deterministic", Get(f, "Kind").ToString());
        foreach (string name in new[] { "", "bad", "1BAD", "A.B", "A B", "A\n", new string('A',32) })
        { f = Formula(); Set(f, "ExcelName", name); Reject(() => Call(f, "Validate")); }
        foreach (string bad in new[] { "", "00000000-0000-0000-0000-000000000000", "{aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa}" })
        { f = Formula(); Set(f, "Id", bad); Reject(() => Call(f, "Validate")); }
        foreach (string field in new[] { "DisplayName", "Description", "OutputUnit" })
        { f = Formula(); Set(f, field, "bad\0text"); Reject(() => Call(f, "Validate")); }
        f = Formula(); Set(f, "DisplayName", " "); Reject(() => Call(f, "Validate"));
        f = Formula(); Set(f, "DisplayName", new string('x',101)); Reject(() => Call(f, "Validate"));
        f = Formula(); Set(f, "Description", new string('x',4001)); Reject(() => Call(f, "Validate"));
        f = Formula(); Set(f, "OutputUnit", new string('x',33)); Reject(() => Call(f, "Validate"));
        f = Formula(); Set(f, "Version", 0); Reject(() => Call(f, "Validate"));
        f = Formula(); Set(f, "Kind", Enum.ToObject(Get(f,"Kind").GetType(),99)); Reject(() => Call(f, "Validate"));
    }
}
