using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Outline.Addin
{
    internal static class AiWorkflowPolicy
    {
        internal static readonly string[] Modes={"scalar","row","column","grid"};
        internal static readonly string[] Routes={"hdec235","qwen72","qwen235","qwen27","deepseek","qwenvl"};
        internal static Dictionary<string,object> RequestInputs(string question,string route)
        {
            if(!Routes.Contains(route))throw new ArgumentException("Unknown workflow route.","route");
            return new Dictionary<string,object>{{"question",question??""},{"route",route}};
        }
        internal static string SerializeProbe(object value,string mode)
        {
            string text;
            if(mode=="scalar")text=Convert.ToString(value)??"";
            else
            {
                var matrix=value as object[,];if(matrix!=null){var rows=new List<string>();int r0=matrix.GetLowerBound(0),c0=matrix.GetLowerBound(1);for(int r=0;r<matrix.GetLength(0);r++){var cells=new List<string>();for(int c=0;c<matrix.GetLength(1);c++)cells.Add(JsonString(matrix[r0+r,c0+c]));rows.Add("["+string.Join(",",cells)+"]");}text="["+string.Join(",",rows)+"]";}
                else {var vector=value as object[];text=vector!=null?"["+string.Join(",",vector.Select(JsonString))+"]":JsonString(value);}
            }
            if(mode=="row")return "Extract requested values and return only one JSON object: {\\\"data\\\":[\\\"value1\\\",\\\"value2\\\"]}. Input: "+text;
            if(mode=="column")return "Extract requested values and return only a JSON object whose data array is a vertical column: {\\\"data\\\":[\\\"value1\\\",\\\"value2\\\"]}. Input: "+text;
            if(mode=="grid")return "Extract requested cells and return only a 2D JSON data array in one object: {\\\"data\\\":[[\\\"r1c1\\\",\\\"r1c2\\\"],[\\\"r2c1\\\",\\\"r2c2\\\"]]}. Keep every row the same width. Input: "+text;
            if(mode!="scalar")throw new ArgumentException("Unknown probe mode.","mode");
            return text;
        }
        static string JsonString(object value){return "\""+(Convert.ToString(value)??"").Replace("\\","\\\\").Replace("\"","\\\"").Replace("\r","\\r").Replace("\n","\\n")+"\"";}
        internal static string CleanAnswer(string answer)
        {
            if(answer==null)return "";
            return Regex.Replace(answer,@"^\s*<think>.*?</think>\s*","",RegexOptions.Singleline|RegexOptions.IgnoreCase);
        }
        internal static object[,] Shape(object value,string mode)
        {
            var wrapped=value as Dictionary<string,object>;
            if(wrapped!=null)
            {
                object data;
                if(!wrapped.TryGetValue("data",out data)||wrapped.Count!=1)throw new ArgumentException("Structured workflow answer must contain only a data field.","value");
                value=data;
            }
            var matrix=value as object[,];
            if(matrix!=null){int r=matrix.GetLength(0),c=matrix.GetLength(1);var copy=new object[r,c];Array.Copy(matrix,copy,matrix.Length);return copy;}
            var vector=value as object[];
            if(vector==null)throw new ArgumentException("Workflow answer must be a JSON array or array of arrays.","value");
            if(mode=="scalar")throw new ArgumentException("Scalar mode requires a scalar workflow answer.","value");
            if(mode=="row"){var row=new object[1,vector.Length];for(int c=0;c<vector.Length;c++)row[0,c]=vector[c];return row;}
            if(mode=="column"){var col=new object[vector.Length,1];for(int r=0;r<vector.Length;r++)col[r,0]=vector[r];return col;}
            if(mode!="grid")throw new ArgumentException("Unknown probe mode.","mode");
            if(vector.Length==0)return new object[0,0];
            var first=vector[0] as object[];if(first==null)throw new ArgumentException("Grid mode requires nested JSON arrays.","value");
            var result=new object[vector.Length,first.Length];
            for(int r=0;r<vector.Length;r++){var current=vector[r] as object[];if(current==null||current.Length!=first.Length)throw new ArgumentException("Grid rows must have equal widths.","value");for(int c=0;c<current.Length;c++)result[r,c]=current[c];}
            return result;
        }
    }
}
