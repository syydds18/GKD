using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Outline.Core;

namespace Outline.Addin
{
    internal static class Ui
    {
        internal static readonly Color Green=Color.FromArgb(33,115,70);
        internal static Button Button(string text, bool primary=false)
        {
            return new Button {FlatStyle=FlatStyle.Flat,UseVisualStyleBackColor=false,Text=text,AutoSize=true,MinimumSize=new Size(100,36),Margin=new Padding(0,4,8,4),
                BackColor=primary?Green:Color.White,ForeColor=primary?Color.White:Color.FromArgb(40,49,59),Padding=new Padding(14,3,14,3),Cursor=Cursors.Hand,
                FlatAppearance={BorderSize=1,BorderColor=primary?Green:Color.FromArgb(212,222,232),MouseOverBackColor=primary?Color.FromArgb(40,132,81):Color.FromArgb(240,245,249),MouseDownBackColor=primary?Color.FromArgb(23,83,50):Color.FromArgb(226,234,241)}};
        }
        internal static void Base(Form form,Size size,string title)
        {
            form.Text=title;form.ClientSize=size;form.MinimumSize=new Size(size.Width+16,size.Height+39);
            form.Font=new Font("Microsoft YaHei UI",9F);form.AutoScaleMode=AutoScaleMode.Dpi;
            form.StartPosition=FormStartPosition.CenterParent;form.BackColor=Color.FromArgb(247,249,251);form.ForeColor=Color.FromArgb(35,47,61);
            form.ShowIcon=false;form.Padding=new Padding(24,20,24,18);
            form.Shown+=(s,e)=>Theme(form);
        }
        private static void Theme(Control root)
        {
            foreach(Control control in root.Controls)
            {
                var grid=control as DataGridView;
                if(grid!=null)
                {
                    grid.BorderStyle=BorderStyle.None;grid.BackgroundColor=Color.White;grid.EnableHeadersVisualStyles=false;
                    grid.CellBorderStyle=DataGridViewCellBorderStyle.SingleHorizontal;grid.GridColor=Color.FromArgb(236,240,244);
                    grid.ColumnHeadersBorderStyle=DataGridViewHeaderBorderStyle.None;grid.ColumnHeadersHeightSizeMode=DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                    grid.ColumnHeadersHeight=42;grid.ColumnHeadersDefaultCellStyle.BackColor=Color.FromArgb(236,241,246);
                    grid.ColumnHeadersDefaultCellStyle.ForeColor=Color.FromArgb(77,92,110);grid.ColumnHeadersDefaultCellStyle.Font=new Font(root.Font,FontStyle.Bold);
                    grid.DefaultCellStyle.Padding=new Padding(8,4,8,4);grid.DefaultCellStyle.SelectionBackColor=Color.FromArgb(214,234,225);
                    grid.DefaultCellStyle.SelectionForeColor=Color.FromArgb(25,71,48);grid.RowTemplate.Height=36;
                    foreach(DataGridViewRow row in grid.Rows)row.Height=36;
                }
                var list=control as CheckedListBox;if(list!=null){list.BorderStyle=BorderStyle.None;list.BackColor=Color.White;list.ItemHeight=30;}
                var combo=control as ComboBox;if(combo!=null)combo.FlatStyle=FlatStyle.Flat;
                var tab=control as TabControl;if(tab!=null){tab.Padding=new Point(18,8);foreach(TabPage page in tab.TabPages){page.BackColor=Color.White;page.Padding=new Padding(1,10,1,1);}}
                Theme(control);
            }
        }
    }
    internal sealed class SheetsForm:Form
    {
        private readonly CheckedListBox list=new CheckedListBox {Dock=DockStyle.Fill,CheckOnClick=true,BorderStyle=BorderStyle.FixedSingle,IntegralHeight=false};
        private readonly List<SheetChoice> choices;
        internal List<SheetChoice> SelectedSheets {get{return list.CheckedIndices.Cast<int>().Select(i=>choices[i]).ToList();}}
        internal SheetsForm(List<SheetChoice> sheets,string active,string action)
        {
            choices=sheets;string verb=action=="check"?"求和校验":action=="clear"?"清除分级":"智能分级";
            Ui.Base(this,new Size(520,430),verb+" — 选择工作表");MaximizeBox=false;
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=4};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,60));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,42));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));
            layout.Controls.Add(new Label {Dock=DockStyle.Fill,Text="选择本次处理范围\r\n直接操作当前工作簿；不会自动保存文件。",AutoSize=false},0,0);
            var toggles=new FlowLayoutPanel {Dock=DockStyle.Fill};
            var all=Ui.Button("全选");var none=Ui.Button("全不选");var current=Ui.Button("仅当前表");
            all.Click+=(s,e)=>SetAll(true);none.Click+=(s,e)=>SetAll(false);current.Click+=(s,e)=>{for(int i=0;i<choices.Count;i++)list.SetItemChecked(i,choices[i].Name==active);};
            toggles.Controls.AddRange(new Control[]{all,none,current});layout.Controls.Add(toggles,0,1);
            foreach(var sheet in choices){int i=list.Items.Add(sheet);if(sheet.Name==active)list.SetItemChecked(i,true);}
            layout.Controls.Add(list,0,2);
            var footer=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft,Padding=new Padding(0,8,0,0)};
            var cancel=Ui.Button("取消");cancel.DialogResult=DialogResult.Cancel;var go=Ui.Button(verb,true);
            go.Click+=(s,e)=>{if(list.CheckedItems.Count==0){MessageBox.Show("请至少选择一张工作表。",verb);return;}DialogResult=DialogResult.OK;};
            footer.Controls.AddRange(new Control[]{cancel,go});layout.Controls.Add(footer,0,3);Controls.Add(layout);AcceptButton=go;CancelButton=cancel;
        }
        private void SetAll(bool value){for(int i=0;i<list.Items.Count;i++)list.SetItemChecked(i,value);}
    }
    internal sealed class ResultRow
    {
        public dynamic Sheet;
        public ValidationError Error;
        public string SheetName {get;private set;}
        public string Address {get{return Error.Letter+Error.Row;}}
        public string Name {get{return Error.RowName;}}
        public string Column {get{return Error.ColLabel;}}
        public string Parent {get{return Error.ParentValue.HasValue?Error.ParentValue.Value.ToString("G12",CultureInfo.InvariantCulture):Error.Message;}}
        public string Sum {get{return Error.ChildSum.HasValue?Error.ChildSum.Value.ToString("G12",CultureInfo.InvariantCulture):"—";}}
        public string Difference {get{return Error.Kind=="sum"?Error.Diff.ToString("G12",CultureInfo.InvariantCulture):"错误值";}}
        public string Basis {get{return Error.Basis;}}
        public string Hint {get{return Error.Hint;}}
        public string Formula {get{return Error.Formula;}}
        internal ResultRow(object sheet,string name,ValidationError error){Sheet=sheet;SheetName=name;Error=error;}
    }
    internal sealed class ResultsForm:Form
    {
        private readonly List<ResultRow> rows;
        private readonly DataGridView grid;
        internal ResultsForm(List<ResultRow> data,int cells,int issues,string notes)
        {
            rows=data;Ui.Base(this,new Size(1080,590),"求和校验结果 — 点击定位，不修改金额");
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=4};
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,67));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,62));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,48));
            var heading=new Label {Dock=DockStyle.Fill,Text="已校验 "+cells+" 个父项金额单元格，发现 "+issues+" 处问题。\r\n原算法采用候选子项择近匹配；默认容差 0.05，不代表完整审计。双击结果可定位原单元格。",ForeColor=issues>0?Color.FromArgb(156,56,32):Ui.Green};layout.Controls.Add(heading,0,0);
            grid=new DataGridView {Dock=DockStyle.Fill,ReadOnly=true,AllowUserToAddRows=false,AllowUserToDeleteRows=false,AutoGenerateColumns=false,
                RowHeadersVisible=false,SelectionMode=DataGridViewSelectionMode.FullRowSelect,MultiSelect=false,BackgroundColor=Color.White,
                BorderStyle=BorderStyle.FixedSingle,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill,DataSource=rows};
            AddColumn("SheetName","工作表",100);AddColumn("Address","位置",65);AddColumn("Name","项目",180);AddColumn("Column","金额列",90);
            AddColumn("Parent","父项金额",90);AddColumn("Sum","子项合计",90);AddColumn("Difference","差额",80);AddColumn("Basis","校验口径",110);
            grid.AlternatingRowsDefaultCellStyle.BackColor=Color.FromArgb(246,248,247);grid.CellDoubleClick+=(s,e)=>Navigate();layout.Controls.Add(grid,0,1);
            var detail=new TextBox {Dock=DockStyle.Fill,ReadOnly=true,Multiline=true,BorderStyle=BorderStyle.None,BackColor=Color.White,Text=notes,ScrollBars=ScrollBars.Vertical};
            grid.SelectionChanged+=(s,e)=>{var row=Current();if(row!=null)detail.Text=(row.Error.Kind=="cell"?"错误："+row.Error.Message:"提示："+(string.IsNullOrEmpty(row.Hint)?"请核对父项公式及子项归属。":row.Hint))+"\r\n公式："+(string.IsNullOrEmpty(row.Formula)?"无公式":row.Formula);};
            layout.Controls.Add(detail,0,2);
            var buttons=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft};
            var close=Ui.Button("关闭");close.Click+=(s,e)=>Close();var locate=Ui.Button("定位单元格",true);locate.Click+=(s,e)=>Navigate();
            var copy=Ui.Button("复制结果");copy.Click+=(s,e)=>Copy();
            buttons.Controls.AddRange(new Control[]{close,locate,copy});layout.Controls.Add(buttons,0,3);Controls.Add(layout);
        }
        private void AddColumn(string property,string caption,int weight){grid.Columns.Add(new DataGridViewTextBoxColumn {DataPropertyName=property,HeaderText=caption,FillWeight=weight,MinimumWidth=55,SortMode=DataGridViewColumnSortMode.NotSortable});}
        private ResultRow Current(){return grid.CurrentRow==null?null:grid.CurrentRow.DataBoundItem as ResultRow;}
        private void Navigate()
        {
            var row=Current();if(row==null)return;
            try {row.Sheet.Parent.Activate();row.Sheet.Activate();ExcelBridge.App.Goto(row.Sheet.Cells[row.Error.Row,row.Error.Col],true);}
            catch(Exception ex){MessageBox.Show("无法定位；工作簿可能已关闭或工作表已变化。\r\n"+ex.Message,"定位失败");}
        }
        private void Copy()
        {
            try
            {
                Func<string,string> safe=s=>{s=(s??"").Replace("\t"," ").Replace("\r"," ").Replace("\n"," ");return s.Length>0&&"=+-@".IndexOf(s[0])>=0?"'"+s:s;};
                string text="工作表\t单元格\t项目\t金额列\t父项金额\t子项合计\t差额\t口径\t提示\r\n"+
                    string.Join("\r\n",rows.Select(r=>string.Join("\t",new[]{r.SheetName,r.Address,r.Name,r.Column,r.Parent,r.Sum,r.Difference,r.Basis,r.Hint}.Select(safe))));
                Clipboard.SetText(text);
            }
            catch(Exception ex){MessageBox.Show(ex.Message,"复制失败");}
        }
    }
}
