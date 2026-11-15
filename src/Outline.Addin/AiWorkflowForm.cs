using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web.Script.Serialization;

namespace Outline.Addin
{
    internal sealed class AiWorkflowForm:Form
    {
        readonly ComboBox mode=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Dock=DockStyle.Fill};
        readonly ComboBox route=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Dock=DockStyle.Fill};
        readonly TextBox prompt=new TextBox{Multiline=true,ScrollBars=ScrollBars.Both,Dock=DockStyle.Fill,AcceptsReturn=true,AcceptsTab=true,WordWrap=false};
        readonly TextBox key=new TextBox{Dock=DockStyle.Fill,UseSystemPasswordChar=true};
        readonly DataGridView results=new DataGridView{Dock=DockStyle.Fill,ReadOnly=true,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill};
        readonly Button run=Ui.Button("运行测试",true);
        readonly Label status=new Label{Dock=DockStyle.Fill,AutoEllipsis=true};
        AiWorkflowConfig config;
        internal AiWorkflowForm()
        {
            Ui.Base(this,new System.Drawing.Size(1000,700),"AI 工作流测试（测试专用）");
            var layout=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=5};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,76));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,76));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,56));
            var settings=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=4,RowCount=2};
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,100));settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50));settings.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,100));settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50));
            settings.Controls.Add(new Label{Text="测试 API Key",Dock=DockStyle.Fill,TextAlign=System.Drawing.ContentAlignment.MiddleLeft},0,0);settings.Controls.Add(key,1,0);
            settings.Controls.Add(new Label{Text="模型路由",Dock=DockStyle.Fill,TextAlign=System.Drawing.ContentAlignment.MiddleLeft},2,0);route.Items.AddRange(AiWorkflowPolicy.Routes);settings.Controls.Add(route,3,0);
            settings.Controls.Add(new Label{Text="Workflow API",Dock=DockStyle.Fill,TextAlign=System.Drawing.ContentAlignment.MiddleLeft},0,1);settings.Controls.Add(new Label{Text="https://aibuilder.hdec.com/v1/workflows/run",Dock=DockStyle.Fill,TextAlign=System.Drawing.ContentAlignment.MiddleLeft},1,1);
            settings.Controls.Add(new Label{Text="用户配置",Dock=DockStyle.Fill,TextAlign=System.Drawing.ContentAlignment.MiddleLeft},2,1);settings.Controls.Add(new Label{Text=AiWorkflowConfig.PathName,Dock=DockStyle.Fill,AutoEllipsis=true,TextAlign=System.Drawing.ContentAlignment.MiddleLeft},3,1);layout.Controls.Add(settings,0,0);
            mode.Items.AddRange(new object[]{"PROMPT scalar","PROMPT row","PROMPT column","PROMPT 2D grid"});mode.SelectedIndex=0;
            var probe=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=2,RowCount=2};probe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,100));probe.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));probe.RowStyles.Add(new RowStyle(SizeType.Absolute,32));probe.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            probe.Controls.Add(new Label{Text="UDF 形状",Dock=DockStyle.Fill},0,0);probe.Controls.Add(mode,1,0);probe.Controls.Add(new Label{Text="PROMPT 输入",Dock=DockStyle.Fill},0,1);probe.Controls.Add(prompt,1,1);layout.Controls.Add(probe,0,1);
            var toolbar=new FlowLayoutPanel{Dock=DockStyle.Fill};var paste=Ui.Button("读取 Excel 选区");paste.Click+=(s,e)=>ReadSelection();run.Click+=(s,e)=>RunProbe();toolbar.Controls.AddRange(new Control[]{paste,run});layout.Controls.Add(toolbar,0,2);
            results.Columns.Add("Mode","探测类型");results.Columns.Add("Route","模型路由");results.Columns.Add("Input","PROMPT 序列化内容");results.Columns.Add("Status","状态");results.Columns.Add("Output","结果 / 错误");layout.Controls.Add(results,0,3);
            layout.Controls.Add(status,0,4);Controls.Add(layout);LoadConfig();
        }
        void LoadConfig(){try{config=AiWorkflowConfig.Load();key.Text=config.ApiKey;int i=Array.IndexOf(AiWorkflowPolicy.Routes,config.User);route.SelectedIndex=i>=0?i:0;}catch(Exception){config=new AiWorkflowConfig();status.Text="读取本地测试配置失败。";}}
        void SaveConfig(){config.ApiKey=key.Text.Trim();config.User=Convert.ToString(route.SelectedItem);config.Save();}
        void ReadSelection()
        {
            try{dynamic selection=ExcelBridge.App.Selection;object value=selection.Value2;int rows=(int)selection.Rows.Count,cols=(int)selection.Columns.Count;
                if(rows==1&&cols==1){prompt.Text=Convert.ToString(value)??"";return;}var matrix=value as object[,];if(matrix==null){prompt.Text=Convert.ToString(value)??"";return;}
                var lines=new System.Collections.Generic.List<string>();for(int r=1;r<=rows;r++){var cells=new System.Collections.Generic.List<string>();for(int c=1;c<=cols;c++)cells.Add(Convert.ToString(matrix[r,c])??"");lines.Add(string.Join("\t",cells));}prompt.Text=string.Join("\r\n",lines);mode.SelectedIndex=rows==1?1:cols==1?2:3;
            }catch(Exception){MessageBox.Show("读取 Excel 选区失败。","AI 测试");}
        }
        async void RunProbe()
        {
            if(!run.Enabled)return;if(string.IsNullOrWhiteSpace(key.Text)){MessageBox.Show("测试 API Key 未配置。","AI 测试");return;}
            try{SaveConfig();}catch(Exception){MessageBox.Show("配置保存失败。","AI 测试");return;}
            string[] modes={"scalar","row","column","grid"};string kind=modes[mode.SelectedIndex],question=AiWorkflowPolicy.SerializeProbe(prompt.Text,kind),selectedRoute=Convert.ToString(route.SelectedItem);
            int index=results.Rows.Add(mode.Text,selectedRoute,question,"运行中…","");run.Enabled=false;status.Text="调用隔离测试工作流…";
            try{string answer=await Task.Run(()=>DifyWorkflowClient.Run(key.Text.Trim(),selectedRoute,question));results.Rows[index].Cells[3].Value="成功";results.Rows[index].Cells[4].Value=answer;status.Text="完成；结果可复制到测试反馈中。";}
            catch(Exception){results.Rows[index].Cells[3].Value="失败";results.Rows[index].Cells[4].Value="请求失败；错误响应已隐藏。";status.Text="调用失败；检查网络和本地测试配置。";}
            finally{run.Enabled=true;}
        }
    }
    internal static class DifyWorkflowClient
    {
        internal static string Run(string apiKey,string route,string question)
        {
            var request=(HttpWebRequest)WebRequest.Create("https://aibuilder.hdec.com/v1/workflows/run");request.Method="POST";request.ContentType="application/json";request.Accept="application/json";request.Headers["Authorization"]="Bearer "+apiKey;request.Timeout=90000;request.ReadWriteTimeout=90000;ServicePointManager.SecurityProtocol|=SecurityProtocolType.Tls12;
            var payload=new System.Collections.Generic.Dictionary<string,object>{{"inputs",AiWorkflowPolicy.RequestInputs(question,route)},{"response_mode","blocking"},{"user","gkd-ai-test"}};
            byte[] body=Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(payload));using(var stream=request.GetRequestStream())stream.Write(body,0,body.Length);
            using(var response=(HttpWebResponse)request.GetResponse())using(var reader=new StreamReader(response.GetResponseStream(),Encoding.UTF8))
            {
                var serializer=new JavaScriptSerializer();var root=serializer.DeserializeObject(reader.ReadToEnd()) as System.Collections.Generic.Dictionary<string,object>;object data,outputs,answer;
                if(root==null||!root.TryGetValue("data",out data)||!(data is System.Collections.Generic.Dictionary<string,object>)||!((System.Collections.Generic.Dictionary<string,object>)data).TryGetValue("outputs",out outputs)||!(outputs is System.Collections.Generic.Dictionary<string,object>)||!((System.Collections.Generic.Dictionary<string,object>)outputs).TryGetValue("answer",out answer))throw new InvalidDataException("Workflow returned no answer.");
                return AiWorkflowPolicy.CleanAnswer(Convert.ToString(answer));
            }
        }
    }
}
