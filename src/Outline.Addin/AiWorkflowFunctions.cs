using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using ExcelDna.Integration;

namespace Outline.Addin
{
    public static class AiWorkflowFunctions
    {
        const string Endpoint="https://aibuilder.hdec.com/v1/workflows/run";
        [ExcelFunction(Name="GKD.PROMPT",Description="Standalone AI workflow scalar probe",IsThreadSafe=true)]
        public static object Prompt(string question,string route){return ExcelAsyncUtil.RunTaskWithCancellation<object>("GKD.PROMPT",new object[]{question,route},ct=>Run(question,route,"scalar",ct));}
        [ExcelFunction(Name="GKD.PROMPTROW",Description="Standalone AI workflow row spill probe",IsThreadSafe=true)]
        public static object PromptRow(object question,string route){return ExcelAsyncUtil.RunTaskWithCancellation<object>("GKD.PROMPTROW",new object[]{question,route},ct=>Run(question,route,"row",ct));}
        [ExcelFunction(Name="GKD.PROMPTCOLUMN",Description="Standalone AI workflow column spill probe",IsThreadSafe=true)]
        public static object PromptColumn(object question,string route){return ExcelAsyncUtil.RunTaskWithCancellation<object>("GKD.PROMPTCOLUMN",new object[]{question,route},ct=>Run(question,route,"column",ct));}
        [ExcelFunction(Name="GKD.PROMPTGRID",Description="Standalone AI workflow 2D spill probe",IsThreadSafe=true)]
        public static object PromptGrid(object question,string route){return ExcelAsyncUtil.RunTaskWithCancellation<object>("GKD.PROMPTGRID",new object[]{question,route},ct=>Run(question,route,"grid",ct));}

        static Task<object> Run(object question,string route,string mode,CancellationToken cancellation)
        {
            return Task.Run<object>(()=>
            {
                cancellation.ThrowIfCancellationRequested();var cfg=AiWorkflowConfig.Load();if(string.IsNullOrWhiteSpace(cfg.ApiKey))return ExcelError.ExcelErrorNA;
                string serialized=AiWorkflowPolicy.SerializeProbe(question,mode);var payload=new Dictionary<string,object>{{"inputs",AiWorkflowPolicy.RequestInputs(serialized,route)},{"response_mode","blocking"},{"user",string.IsNullOrWhiteSpace(cfg.User)?"gkd-excel-udf":cfg.User}};
                var request=(HttpWebRequest)WebRequest.Create(Endpoint);request.Method="POST";request.ContentType="application/json";request.Headers["Authorization"]="Bearer "+cfg.ApiKey;request.Timeout=90000;request.ReadWriteTimeout=90000;
                using(cancellation.Register(()=>{try{request.Abort();}catch{}}))
                {
                    byte[] bytes=Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(payload));using(var stream=request.GetRequestStream())stream.Write(bytes,0,bytes.Length);
                    using(var response=(HttpWebResponse)request.GetResponse())using(var reader=new System.IO.StreamReader(response.GetResponseStream(),Encoding.UTF8))
                    {
                        var root=new JavaScriptSerializer().DeserializeObject(reader.ReadToEnd()) as Dictionary<string,object>;object data,outputs,answer;
                        if(root==null||!root.TryGetValue("data",out data)||!(data is Dictionary<string,object>)||!((Dictionary<string,object>)data).TryGetValue("outputs",out outputs)||!(outputs is Dictionary<string,object>)||!((Dictionary<string,object>)outputs).TryGetValue("answer",out answer))return ExcelError.ExcelErrorValue;
                        string text=AiWorkflowPolicy.CleanAnswer(Convert.ToString(answer));if(mode=="scalar")return (object)text;
                        object parsed;try{parsed=new JavaScriptSerializer().DeserializeObject(text);}catch{return ExcelError.ExcelErrorValue;}
                        try{return (object)AiWorkflowPolicy.Shape(parsed,mode);}catch{return ExcelError.ExcelErrorValue;}
                    }
                }
            },cancellation);
        }
    }
}
