using System;
using System.Collections.Generic;
using Outline.Addin;

internal static class Program
{
    static int checks;
    static void Check(bool ok,string label){if(!ok)throw new Exception(label);checks++;}
    static void Main()
    {
        Check((string)AiWorkflowPolicy.RequestInputs("hello","qwen72")["question"]=="hello","confirmed question/route schema");
        Check(AiWorkflowPolicy.SerializeProbe(new object[,]{{"a","b"},{"c","d"}},"grid").EndsWith("[[\"a\",\"b\"],[\"c\",\"d\"]]"),"grid request serialization");
        Check(AiWorkflowPolicy.CleanAnswer("<think>hidden</think>\n[1]")=="[1]","strip leading think block");
        var vector=new object[]{"a","b"};
        var wrapped=AiWorkflowPolicy.Shape(new Dictionary<string,object>{{"data",vector}},"row");
        Check(wrapped.GetLength(0)==1&&wrapped.GetLength(1)==2&&(string)wrapped[0,1]=="b","unwrap workflow data object for row spill");
        Check(AiWorkflowPolicy.SerializeProbe("sample","row").Contains("JSON object"),"row mode includes structured-output instruction");
        Check(AiWorkflowPolicy.SerializeProbe("sample","column").Contains("column"),"column mode includes orientation instruction");
        Check(AiWorkflowPolicy.SerializeProbe("sample","grid").Contains("2D JSON"),"grid mode includes matrix instruction");
        var shaped=AiWorkflowPolicy.Shape(vector,"row");
        Check(shaped.GetLength(0)==1&&shaped.GetLength(1)==2&&(string)shaped[0,1]=="b","row spill");
        shaped=AiWorkflowPolicy.Shape(vector,"column");
        Check(shaped.GetLength(0)==2&&shaped.GetLength(1)==1,"column spill");
        shaped=AiWorkflowPolicy.Shape(new object[]{new object[]{"a","b"},new object[]{"c","d"}},"grid");
        Check(shaped.GetLength(0)==2&&(string)shaped[1,1]=="d","grid spill");
        bool ragged=false;try{AiWorkflowPolicy.Shape(new object[]{new object[]{"a"},new object[]{"b","c"}},"grid");}catch(ArgumentException){ragged=true;}
        Check(ragged,"reject ragged matrix");
        bool route=false;try{AiWorkflowPolicy.RequestInputs("x","invented");}catch(ArgumentException){route=true;}
        Check(route,"reject unknown route");
        Console.WriteLine("PASS: "+checks+" AI workflow policy checks");
    }
}
