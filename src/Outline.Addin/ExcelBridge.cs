using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using Outline.Core;

namespace Outline.Addin
{
    public sealed class AddinEntry : IExcelAddIn
    {
        public void AutoOpen() { GkdCustomFormulaFunctions.Start(); GkdAiTableFormatting.Start(); GkdAiWelcome.Start(); }
        public void AutoClose() { GkdCustomFormulaFunctions.Stop(); GkdAiWelcome.Stop(); GkdAiTableFormatting.Stop(); Controller.Close(); }
    }
    internal sealed class SheetSnapshot
    {
        public dynamic Sheet;
        public int MaxRow, MaxCol, Summary;
        public string DataFingerprint;
        public object Workbook;
        public int[] RowLevels, ColLevels;
        public bool[] RowHidden, ColHidden;
    }
    internal static class ExcelBridge
    {
        public static dynamic App { get { return ExcelDnaUtil.Application; } }
        public static string ErrorText(object value)
        {
            var e = value as ErrorWrapper;
            if (e == null) return null;
            int code = e.ErrorCode & 0xffff;
            switch(code) { case 2000:return "#NULL!"; case 2007:return "#DIV/0!"; case 2015:return "#VALUE!"; case 2023:return "#REF!"; case 2029:return "#NAME?"; case 2036:return "#NUM!"; case 2042:return "#N/A"; case 2045:return "#SPILL!"; case 2050:return "#CALC!"; default:return "#VALUE!"; }
        }
        public static SheetData Read(dynamic sheet)
        {
            dynamic used=sheet.UsedRange;
            int lastRow=(int)used.Row+(int)used.Rows.Count-1;
            int lastCol=(int)used.Column+(int)used.Columns.Count-1;
            if ((long)lastRow*lastCol>2000000L) throw new InvalidOperationException("工作表使用区域过大（超过 200 万个单元格），请先清理远端空白格式后重试。没有修改该表。");
            dynamic range=sheet.Range[sheet.Cells[1,1],sheet.Cells[lastRow,lastCol]];
            object values=range.Value2; object formulas=range.Formula;
            var a=values as object[,]; var f=formulas as object[,];
            var data=new SheetData { MaxRow=lastRow, MaxCol=lastCol };
            for(int r=1;r<=lastRow;r++)
            {
                var row=new Dictionary<int,object>(); var fr=new Dictionary<int,string>();
                for(int c=1;c<=lastCol;c++)
                {
                    object v=a==null?values:a[r,c];
                    var err=ErrorText(v); if(err!=null)v=err;
                    if(v!=null && Convert.ToString(v,CultureInfo.InvariantCulture)!="") row[c]=v;
                    object fv=f==null?formulas:f[r,c]; var fs=fv as string;
                    if(fs!=null && fs.StartsWith("=",StringComparison.Ordinal))fr[c]=fs.Substring(1);
                }
                if(row.Count>0)data.Values[r]=row;
                if(fr.Count>0)data.Formulas[r]=fr;
                int lv=Math.Max(0,(int)sheet.Rows[r].OutlineLevel-1);
                data.OutlineLevels[r]=lv;
                data.OutlineMax=Math.Max(data.OutlineMax,lv);
            }
            return data;
        }
        public static SheetSnapshot Capture(dynamic sheet,int rowCount,int colCount)
        {
            var s=new SheetSnapshot {Sheet=sheet,MaxRow=rowCount,MaxCol=colCount,Summary=(int)sheet.Outline.SummaryRow,
                RowLevels=new int[rowCount+1],RowHidden=new bool[rowCount+1],ColLevels=new int[colCount+1],ColHidden=new bool[colCount+1]};
            for(int r=1;r<=rowCount;r++){s.RowLevels[r]=Math.Max(0,(int)sheet.Rows[r].OutlineLevel-1);s.RowHidden[r]=(bool)sheet.Rows[r].Hidden;}
            for(int c=1;c<=colCount;c++){s.ColLevels[c]=Math.Max(0,(int)sheet.Columns[c].OutlineLevel-1);s.ColHidden[c]=(bool)sheet.Columns[c].Hidden;}
            return s;
        }
        private static void Groups(dynamic sheet,int[] levels,bool columns)
        {
            int n=levels.Length-1;
            for(int depth=1;depth<=7;depth++)
            {
                int begin=0;
                for(int i=1;i<=n+1;i++)
                {
                    bool inside=i<=n && levels[i]>=depth;
                    if(inside && begin==0)begin=i;
                    if(!inside && begin>0)
                    {
                        if(columns)sheet.Range[sheet.Cells[1,begin],sheet.Cells[1,i-1]].EntireColumn.Group();
                        else sheet.Rows[begin.ToString(CultureInfo.InvariantCulture)+":"+(i-1).ToString(CultureInfo.InvariantCulture)].Group();
                        begin=0;
                    }
                }
            }
        }
        private static void SetHidden(dynamic sheet,bool[] hidden,bool columns)
        {
            int n=hidden.Length-1;int begin=1;
            while(begin<=n)
            {
                bool state=hidden[begin];int end=begin;
                while(end<n && hidden[end+1]==state)end++;
                if(columns)sheet.Range[sheet.Cells[1,begin],sheet.Cells[1,end]].EntireColumn.Hidden=state;
                else sheet.Rows[begin+":"+end].Hidden=state;
                begin=end+1;
            }
        }
        private static dynamic Scope(dynamic sheet,int rows,int cols)
        {
            return sheet.Range[sheet.Cells[1,1],sheet.Cells[rows,cols]];
        }
        public static void Apply(dynamic sheet,Detection detection,SheetSnapshot snapshot,bool clear)
        {
            Scope(sheet,snapshot.MaxRow,snapshot.MaxCol).ClearOutline();
            // Original tool clears old row/column outline and hidden states in its recorded extent.
            sheet.Rows["1:"+snapshot.MaxRow].Hidden=false;
            sheet.Range[sheet.Cells[1,1],sheet.Cells[1,snapshot.MaxCol]].EntireColumn.Hidden=false;
            sheet.Outline.SummaryRow=0; // xlAbove: subtotal row precedes children
            if(!clear)
            {
                var levels=new int[snapshot.MaxRow+1];
                foreach(var pair in OutlineHierarchy.Compact(detection.Levels))if(pair.Key>0 && pair.Key<levels.Length)levels[pair.Key]=Math.Max(0,Math.Min(7,pair.Value));
                Groups(sheet,levels,false);
            }
        }
        public static void Restore(SheetSnapshot snapshot)
        {
            dynamic sheet=snapshot.Sheet;
            Scope(sheet,snapshot.MaxRow,snapshot.MaxCol).ClearOutline();
            Groups(sheet,snapshot.RowLevels,false);Groups(sheet,snapshot.ColLevels,true);
            sheet.Outline.SummaryRow=snapshot.Summary;
            SetHidden(sheet,snapshot.RowHidden,false);SetHidden(sheet,snapshot.ColHidden,true);
        }
    }
    internal sealed class SheetChoice
    {
        public dynamic Sheet;
        public string Name;
        public bool IsProtected;
        public override string ToString(){return Name+(IsProtected?"（受保护）":"");}
    }
    internal static class Controller
    {
        private static readonly List<SheetSnapshot> Undo=new List<SheetSnapshot>();
        private static ResultsForm results;
        private static GkdCustomFormulaForm customFormulaForm;
        public static void Close(){if(customFormulaForm!=null)customFormulaForm.CloseForShutdown();GkdAiFunctions.Scheduler.CancelAll();SmartSumController.ClosePreview();if(results!=null)results.Close();Undo.Clear();}
        public static void Execute(string action)
        {
            try
            {
                if(action=="customFormula")
                {
                    if(customFormulaForm!=null && !customFormulaForm.IsDisposed){customFormulaForm.Activate();return;}
                    using(var form=new GkdCustomFormulaForm())
                    {
                        customFormulaForm=form;
                        try{form.ShowDialog(new ExcelOwner());}
                        finally{customFormulaForm=null;}
                    }
                    return;
                }
                if(action=="aiSettings"){using(var form=new GkdAiSettingsForm())form.ShowDialog(new ExcelOwner());return;}
                if(action=="aiRefresh" || action=="aiCancel" || action=="aiValues"){GkdAiActions.Run(action);return;}
                if(action=="help"){using(var form=new GeneralHelpForm())form.ShowDialog(new ExcelOwner());return;}
                dynamic app=ExcelBridge.App;
                if(app.ActiveWorkbook==null){MessageBox.Show("请先打开需要处理的工作簿。","搞快点");return;}
                if(action=="smartSum"){SmartSumController.Run();return;}
                if(action=="projectExtract"){ProjectExtractionRunner.Run();return;}
                if(action=="undoSum"){SmartSumController.Undo();return;}
                if(action=="restore"){Restore();return;}
                if(action=="expand" || action=="level1" || action=="level2" || action=="level3")
                {
                    dynamic ws=app.ActiveSheet;
                    ws.Outline.ShowLevels(action=="expand"?8:int.Parse(action.Substring(5),CultureInfo.InvariantCulture),Type.Missing);return;
                }
                var choices=new List<SheetChoice>();
                foreach(dynamic ws in app.ActiveWorkbook.Worksheets)choices.Add(new SheetChoice {Sheet=ws,Name=(string)ws.Name,IsProtected=(bool)ws.ProtectContents});
                using(var chooser=new SheetsForm(choices,(string)app.ActiveSheet.Name,action))
                {
                    if(chooser.ShowDialog(new ExcelOwner())!=DialogResult.OK)return;
                    var selected=chooser.SelectedSheets;
                    if(action=="check"){Check(selected);return;}
                    bool clear=action=="clear";
                    if(MessageBox.Show((clear?"清除":"重建")+"所选 "+selected.Count+" 张工作表的分级？\n\n会清除原有行列分组，并展开已隐藏的行列。数值、公式、样式、合并均不改。\n操作后可使用“恢复上次分级”。","确认操作",MessageBoxButtons.OKCancel,MessageBoxIcon.Warning)!=DialogResult.OK)return;
                    Mutate(selected,clear);
                }
            }
            catch(Exception ex){MessageBox.Show(ex.Message,"搞快点 — 未完成",MessageBoxButtons.OK,MessageBoxIcon.Error);}
        }
        private static void Mutate(List<SheetChoice> selected,bool clear)
        {
            var planned=new List<Tuple<SheetSnapshot,Detection>>();
            foreach(var c in selected)
            {
                if(c.IsProtected)throw new InvalidOperationException("“"+c.Name+"”受保护，请先取消工作表保护。没有开始修改。");
                if((bool)c.Sheet.FilterMode)throw new InvalidOperationException("“"+c.Name+"”正在筛选，为避免改变筛选隐藏状态，请先清除筛选条件。没有开始修改。");
                var data=ExcelBridge.Read((object)c.Sheet);var det=Engine.Detect(data);
                if(!clear && det.MaxLevel<=0)throw new InvalidOperationException("“"+c.Name+"”未识别到可用编号层级，没有开始修改。请检查表头“编号/序号”和“名称”。");
                var snapshot=ExcelBridge.Capture((object)c.Sheet,data.MaxRow,data.MaxCol);
                snapshot.DataFingerprint=SnapshotGuard.Fingerprint(data);snapshot.Workbook=c.Sheet.Parent;
                planned.Add(Tuple.Create(snapshot,det));
            }
            dynamic app=ExcelBridge.App;bool screen=app.ScreenUpdating;bool events=app.EnableEvents;var touched=new List<SheetSnapshot>();
            try
            {
                app.ScreenUpdating=false;app.EnableEvents=false;
                foreach(var plan in planned){touched.Add(plan.Item1);ExcelBridge.Apply(plan.Item1.Sheet,plan.Item2,plan.Item1,clear);}
                Undo.Clear();Undo.AddRange(touched);
            }
            catch(Exception ex)
            {
                var failures=new List<string>();foreach(var snap in touched){try{ExcelBridge.Restore(snap);}catch(Exception r){failures.Add(r.Message);}}
                throw new InvalidOperationException(ex.Message+(failures.Count==0?"\n本批已回退原分级。":"\n部分回退失败，请勿保存并从原文件重新打开："+string.Join("; ",failures)));
            }
            finally{app.ScreenUpdating=screen;app.EnableEvents=events;app.StatusBar=false;}
            MessageBox.Show("已"+(clear?"清除":"重建")+" "+touched.Count+" 张工作表的分级。\n工作簿尚未自动保存。","操作完成");
        }
        private static void Restore()
        {
            if(Undo.Count==0){MessageBox.Show("当前会话没有可恢复的分级快照。","恢复上次分级");return;}
            dynamic app=ExcelBridge.App;
            foreach(var snap in Undo)
            {
                if(!Object.ReferenceEquals((object)app.ActiveWorkbook,snap.Workbook))
                    throw new InvalidOperationException("请切回执行上次分级的工作簿，再恢复；没有修改当前工作簿。");
                var current=ExcelBridge.Read(snap.Sheet);
                if(SnapshotGuard.Fingerprint(current)!=snap.DataFingerprint)
                    throw new InvalidOperationException("数据或工作表范围已变化，为避免按旧行号恢复到错误位置，已拒绝恢复。请从保存的原文件重新打开。");
                if((bool)snap.Sheet.ProtectContents)throw new InvalidOperationException("工作表已受保护，请先取消保护再恢复。");
                if((bool)snap.Sheet.FilterMode)throw new InvalidOperationException("工作表正在筛选，请先清除筛选条件再恢复。");
            }
            if(MessageBox.Show("恢复最近一批分级及隐藏状态？\n仅恢复分级及隐藏状态，不改金额和公式。", "恢复上次分级",MessageBoxButtons.OKCancel,MessageBoxIcon.Warning)!=DialogResult.OK)return;
            bool screen=app.ScreenUpdating;bool events=app.EnableEvents;
            try{app.ScreenUpdating=false;app.EnableEvents=false;foreach(var snap in Undo)ExcelBridge.Restore(snap);Undo.Clear();}
            finally{app.ScreenUpdating=screen;app.EnableEvents=events;}
        }
        private static void Check(List<SheetChoice> choices)
        {
            var rows=new List<ResultRow>();int cells=0,issues=0;var notes=new List<string>();
            foreach(var c in choices)
            {
                var data=ExcelBridge.Read((object)c.Sheet);var det=Engine.Detect(data);var validation=Engine.Validate(data,det);
                cells+=validation.CheckedCells;issues+=validation.ErrorCount;
                if(validation.Status=="na")notes.Add(c.Name+"："+validation.Reason);
                foreach(var error in validation.Errors)rows.Add(new ResultRow(c.Sheet,c.Name,error));
                if(validation.Truncated)notes.Add(c.Name+"：还有 "+validation.HiddenErrors+" 条未展开（原算法每表最多显示 500 条）。");
            }
            if(results!=null && !results.IsDisposed)results.Close();
            results=new ResultsForm(rows,cells,issues,string.Join("\r\n",notes));results.Show(new ExcelOwner());
        }
    }
    internal sealed class ExcelOwner : IWin32Window { public IntPtr Handle { get { return new IntPtr((int)ExcelBridge.App.Hwnd); } } }
}
