using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace Outline.Addin
{
    internal sealed class AiWorkflowConfig
    {
        public string Endpoint="https://aibuilder.hdec.com/v1/workflows/run";
        public string ApiKey="";
        public string User="hdec235";
        public Dictionary<string,string> Aliases=new Dictionary<string,string>();
        internal static string PathName {get{return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"GKD","ai-workflow.json");}}
        internal static AiWorkflowConfig Load()
        {
            if(!File.Exists(PathName))return new AiWorkflowConfig();
            var cfg=new JavaScriptSerializer().Deserialize<AiWorkflowConfig>(File.ReadAllText(PathName,Encoding.UTF8));
            if(cfg==null)throw new InvalidDataException("配置文件内容无效。");
            if(cfg.Aliases==null)cfg.Aliases=new Dictionary<string,string>();
            return cfg;
        }
        internal void Save()
        {
            string dir=System.IO.Path.GetDirectoryName(PathName);Directory.CreateDirectory(dir);
            File.WriteAllText(PathName,new JavaScriptSerializer().Serialize(this),new UTF8Encoding(false));
        }
        internal static string Template(){return new JavaScriptSerializer().Serialize(new AiWorkflowConfig());}
    }
}
