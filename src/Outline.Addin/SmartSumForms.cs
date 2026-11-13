using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Outline.Core;

namespace Outline.Addin
{
    internal sealed class SumSettingsForm:Form
    {
        private readonly CheckedListBox columns=new CheckedListBox {Dock=DockStyle.Fill,CheckOnClick=true,IntegralHeight=false};
        private readonly ComboBox source=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Dock=DockStyle.Fill};
        internal List<int> Columns {get{return columns.CheckedIndices.Cast<int>().Select(i=>i+1).ToList();}}
        internal bool UseOutline {get{return source.SelectedIndex==1;}}
        private readonly NumericUpDown startRow=new NumericUpDown {Minimum=1,Maximum=1048576,Width=90};
        private readonly ComboBox nameColumn=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Width=85};
        internal Detection OutlineLayout {get{return new Detection {HeaderRow=(int)startRow.Value-1,NameCol=nameColumn.SelectedIndex+1,CodeCol=0};}}
        internal SumSettingsForm(SheetData data,Detection detection,string sheetName)
        {
            Ui.Base(this,new Size(660,620),"智能求和 — "+sheetName);
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=7};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,62));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,38));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,76));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,55));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));
            layout.Controls.Add(new Label {Dock=DockStyle.Fill,Text="先指定列，再逐项预览确认。仅处理当前工作表。\r\n按直接子项写 SUM，不把中间汇总与下级明细重复相加。"},0,0);
            layout.Controls.Add(new Label {Dock=DockStyle.Fill,Text="父子关系来源（不会自动切换）"},0,1);
            source.Items.AddRange(new object[]{"按编号识别（推荐）","使用现有 Excel 行分级（汇总行必须在上）"});source.SelectedIndex=0;
            layout.Controls.Add(source,0,2);
            var outlineFields=new FlowLayoutPanel {Dock=DockStyle.Fill,Enabled=false};
            var outlineLayout=SmartSumPlanner.FindOutlineLayout(data);
            startRow.Maximum=Math.Max(1,data.MaxRow);
            startRow.Value=Math.Max(1,Math.Min(data.MaxRow,(outlineLayout??detection).HeaderRow+1));
            for(int c=1;c<=data.MaxCol;c++)nameColumn.Items.Add(Col(c));
            if(nameColumn.Items.Count>0)nameColumn.SelectedIndex=Math.Max(0,Math.Min(data.MaxCol-1,(outlineLayout??detection).NameCol-1));
            outlineFields.Controls.AddRange(new Control[]{new Label {Text="数据起始行",AutoSize=true},startRow,new Label {Text="名称列",AutoSize=true},nameColumn,
                new Label {Text="仅原生分级模式：不要求编号；请确认起始行包含最上层父项。",AutoSize=true,ForeColor=Ui.Green}});
            source.SelectedIndexChanged+=(s,e)=>outlineFields.Enabled=UseOutline;
            layout.Controls.Add(outlineFields,0,3);
            var picks=new FlowLayoutPanel {Dock=DockStyle.Fill};var recommended=Ui.Button("推荐金额列");var none=Ui.Button("清空选择");
            var suggested=new List<int>();
            for(int c=1;c<=data.MaxCol;c++)
            {
                string label="";Dictionary<int,object> row;object val;
                if(data.Values.TryGetValue((outlineLayout??detection).HeaderRow,out row)&&row.TryGetValue(c,out val))label=Convert.ToString(val);
                bool amount=Engine.IsAmountHeaderLabel(label??"");if(amount)suggested.Add(c);
                columns.Items.Add(Col(c)+" · "+(string.IsNullOrWhiteSpace(label)?"（无表头）":label)+(amount?"  [推荐]":""),amount);
            }
            recommended.Click+=(s,e)=>{for(int i=0;i<columns.Items.Count;i++)columns.SetItemChecked(i,suggested.Contains(i+1));};
            none.Click+=(s,e)=>{for(int i=0;i<columns.Items.Count;i++)columns.SetItemChecked(i,false);};
            picks.Controls.AddRange(new Control[]{recommended,none});layout.Controls.Add(picks,0,4);layout.Controls.Add(columns,0,5);
            var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft};
            var cancel=Ui.Button("取消");cancel.DialogResult=DialogResult.Cancel;var preview=Ui.Button("生成预览",true);
            preview.Click+=(s,e)=>{if(Columns.Count==0){MessageBox.Show("请至少选择一列。","指定列");return;}var selectedLayout=UseOutline?OutlineLayout:detection;if(Columns.Contains(selectedLayout.NameCol)||!UseOutline&&Columns.Contains(selectedLayout.CodeCol)){MessageBox.Show(UseOutline?"名称列不能作为求和写入列，请取消勾选。":"编号列和名称列不能作为求和写入列，请取消勾选。","保护表结构");return;}DialogResult=DialogResult.OK;};
            buttons.Controls.AddRange(new Control[]{cancel,preview});layout.Controls.Add(buttons,0,6);Controls.Add(layout);CancelButton=cancel;AcceptButton=preview;
        }
        internal static string Col(int col){var s="";while(col>0){col--;s=(char)('A'+col%26)+s;col/=26;}return s;}
    }
    internal sealed class SumPreviewForm:Form
    {
        private readonly DataGridView grid;
        private readonly List<SumProposal> proposals;
        private readonly Label summary=new Label {Dock=DockStyle.Fill,AutoEllipsis=true};
        private readonly Action<int,int> navigate;
        private readonly Dictionary<string,List<string>> dependencies=new Dictionary<string,List<string>>();
        internal event Action SubmitRequested;
        internal List<SumProposal> Selected {get;private set;}
        internal SumPreviewForm(SumPlan plan,SheetData data,Action<int,int> navigate)
        {
            this.navigate=navigate;proposals=plan.Proposals.OrderBy(p=>p.Row).ThenBy(p=>p.Col).ToList();
            Ui.Base(this,new Size(1200,700),"智能求和 — 逐项确认");
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=5};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,46));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,70));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,38));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));
            layout.Controls.Add(new Label {Dock=DockStyle.Fill,Text="默认全部不勾选。双击定位至工作表视野中部；可滚动原 Excel 页面，确认窗口保持打开。\r\n源数据或分级变化后请重新生成预览；全选会包含已有数值和公式，写入前仍需最终确认。"},0,0);
            grid=NewGrid();grid.ShowCellToolTips=true;
            grid.Columns.Add(new DataGridViewCheckBoxColumn {Name="Pick",HeaderText="确认",FillWeight=32,MinimumWidth=48,SortMode=DataGridViewColumnSortMode.NotSortable});
            foreach(var c in new[]{new[]{"State","情况"},new[]{"Address","位置"},new[]{"Name","项目"},new[]{"Before","原值 / 原公式"},new[]{"After","新公式"},new[]{"Children","直接子项行"}})
                grid.Columns.Add(new DataGridViewTextBoxColumn {Name=c[0],HeaderText=c[1],ReadOnly=true,FillWeight=c[0]=="After"?190:c[0]=="Address"?60:110,MinimumWidth=65,SortMode=DataGridViewColumnSortMode.NotSortable});
            var byAddress=proposals.ToDictionary(p=>p.Address,StringComparer.Ordinal);
            foreach(var p in proposals)
            {
                var required=p.ChildRows.Select(r=>PreviewPolicy.Col(p.Col)+r).Where(a=>byAddress.ContainsKey(a)&&PreviewPolicy.IsBlank(byAddress[a])).ToList();
                dependencies[p.Address]=required;
                string state=PreviewPolicy.Kind(p)+(required.Count>0?" · 依赖子汇总":"");
                string old=p.HasExistingFormula&&!p.OriginalText.StartsWith("=",StringComparison.Ordinal)?"="+p.OriginalText:p.OriginalText;
                int i=grid.Rows.Add(false,state,p.Address,p.RowName,old,p.Formula,string.Join(", ",p.ChildRows));
                var row=grid.Rows[i];row.DefaultCellStyle.BackColor=StateColor(p);
                string tip=PreviewPolicy.RowPreview(data,p.Row)+"\r\n新公式："+p.Formula+
                    (required.Count>0?"\r\n须同时勾选空白子汇总："+string.Join(", ",required):"");
                foreach(DataGridViewCell cell in row.Cells)cell.ToolTipText=tip;
            }
            grid.CurrentCellDirtyStateChanged+=(s,e)=>{if(grid.IsCurrentCellDirty)grid.CommitEdit(DataGridViewDataErrorContexts.Commit);};
            grid.CellValueChanged+=(s,e)=>UpdateSummary();
            grid.CellDoubleClick+=(s,e)=>{if(e.RowIndex>=0&&e.ColumnIndex>0)Locate(proposals[e.RowIndex].Row,proposals[e.RowIndex].Col);};
            var toolbar=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoScroll=true};
            var all=Ui.Button("全选");all.Click+=(s,e)=>SetChecks(false,true);
            var none=Ui.Button("取消全选");none.Click+=(s,e)=>SetChecks(false,false);
            var blanks=Ui.Button("仅勾选原值为空");blanks.Click+=(s,e)=>SetChecks(true,true);
            var next=Ui.Button("下一个已有值/公式");next.Width=175;next.Click+=(s,e)=>{
                int i=PreviewPolicy.NextExisting(proposals,grid.CurrentCell==null?-1:grid.CurrentCell.RowIndex);
                if(i<0){MessageBox.Show(this,"所有候选目标原值均为空。","没有已有内容");return;}
                grid.CurrentCell=grid.Rows[i].Cells["Address"];grid.FirstDisplayedScrollingRowIndex=i;
                Locate(proposals[i].Row,proposals[i].Col);
            };
            toolbar.Controls.AddRange(new Control[]{all,none,blanks,next});layout.Controls.Add(toolbar,0,1);
            var tabs=new TabControl {Dock=DockStyle.Fill};var candidates=new TabPage("求和候选");candidates.Controls.Add(grid);tabs.TabPages.Add(candidates);
            if(plan.Warnings.Count>0)
            {
                var skipped=new TabPage("跳过 / 警告（"+plan.Warnings.Count+"）");var warnings=NewGrid();warnings.ReadOnly=true;
                warnings.Columns.Add("State","情况");warnings.Columns.Add("Reason","原因（未生成公式，不参与勾选）");warnings.Columns[0].FillWeight=25;warnings.Columns[1].FillWeight=200;
                foreach(string warning in plan.Warnings){int n=warnings.Rows.Add("跳过 / 需复核",warning);warnings.Rows[n].DefaultCellStyle.BackColor=Color.FromArgb(255,235,235);warnings.Rows[n].Cells[1].ToolTipText=warning;}
                skipped.Controls.Add(warnings);tabs.TabPages.Add(skipped);
            }
            layout.Controls.Add(tabs,0,2);layout.Controls.Add(summary,0,3);
            var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft};
            var cancel=Ui.Button("取消");cancel.Click+=(s,e)=>Close();var apply=Ui.Button("写入勾选项",true);
            apply.Click+=(s,e)=>{
                grid.EndEdit();Selected=new List<SumProposal>();for(int i=0;i<grid.Rows.Count;i++)if(Convert.ToBoolean(grid.Rows[i].Cells[0].Value))Selected.Add(proposals[i]);
                if(Selected.Count==0){MessageBox.Show(this,"请勾选需要写入的公式。","尚未确认");return;}
                var chosen=new HashSet<string>(Selected.Select(p=>p.Address),StringComparer.Ordinal);
                var missing=Selected.SelectMany(p=>dependencies[p.Address].Where(a=>!chosen.Contains(a)).Select(a=>p.Address+" → "+a)).ToList();
                if(missing.Count>0){MessageBox.Show(this,"以下父项依赖尚未勾选的空白子汇总，请同时勾选：\r\n"+string.Join("\r\n",missing.Take(15)),"避免漏计");return;}
                if(SubmitRequested!=null)SubmitRequested();
            };
            buttons.Controls.AddRange(new Control[]{cancel,apply});layout.Controls.Add(buttons,0,4);Controls.Add(layout);CancelButton=cancel;UpdateSummary();
        }
        private static DataGridView NewGrid(){return new DataGridView {Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false,BackgroundColor=Color.White,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill,SelectionMode=DataGridViewSelectionMode.FullRowSelect,MultiSelect=false};}
        private static Color StateColor(SumProposal p){return p.HasExistingFormula?Color.FromArgb(230,239,255):PreviewPolicy.IsBlank(p)?Color.FromArgb(233,247,238):Color.FromArgb(255,243,218);}
        private bool updatingChecks;
        private void SetChecks(bool onlyBlank,bool value)
        {
            grid.EndEdit();var checks=PreviewPolicy.Selection(proposals,onlyBlank?"blank":value?"all":"none");
            updatingChecks=true;try{for(int i=0;i<checks.Length;i++)grid.Rows[i].Cells[0].Value=checks[i];}finally{updatingChecks=false;UpdateSummary();}
        }
        private void UpdateSummary()
        {
            if(grid==null||updatingChecks)return;int count=0;foreach(DataGridViewRow row in grid.Rows)if(Convert.ToBoolean(row.Cells[0].Value))count++;
            summary.Text="已勾选 "+count+" / "+proposals.Count+"　图例：浅绿＝原值为空　浅黄＝已有数值　浅蓝＝已有公式　浅红＝跳过/警告";
        }
        private void Locate(int row,int col){try{navigate(row,col);}catch(Exception ex){MessageBox.Show(this,"定位失败："+ex.Message,"Excel 定位");}}
    }
}
