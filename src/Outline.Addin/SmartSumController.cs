using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Outline.Core;

namespace Outline.Addin
{
    internal static class SmartSumController
    {
        private static object undoSheet, undoWorkbook;
        private static string undoFingerprint;
        private static List<FormulaEdit> undo;
        internal static void Run()
        {
            dynamic app=ExcelBridge.App;
            if(app.ActiveWorkbook==null)throw new InvalidOperationException("请先打开工作簿。");
            dynamic sheet=app.ActiveSheet;
            if((int)sheet.Type!=-4167)throw new InvalidOperationException("请先切换到普通工作表。");
            if((bool)sheet.ProtectContents)throw new InvalidOperationException("请先取消工作表保护。");
            if((bool)sheet.FilterMode)throw new InvalidOperationException("请先清除筛选条件，避免误判隐藏项归属。");
            var data=ExcelBridge.Read((object)sheet);var detection=Engine.Detect(data);
            string before=SnapshotGuard.Fingerprint(data);
            using(var settings=new SumSettingsForm(data,detection,(string)sheet.Name))
            {
                if(settings.ShowDialog(new ExcelOwner())!=DialogResult.OK)return;
                if(settings.UseOutline && (int)sheet.Outline.SummaryRow!=0)throw new InvalidOperationException("当前原生分级的汇总行在下方；智能求和仅支持父项在上。请改用编号识别并复核预览。");
                var plan=SmartSumPlanner.Build(data,settings.UseOutline?settings.OutlineLayout:detection,settings.Columns,settings.UseOutline);
                if(plan.Proposals.Count==0){MessageBox.Show("没有可安全生成的公式。\r\n"+string.Join("\r\n",plan.Warnings.Take(15)),"智能求和");return;}
                ClosePreview();
                var preview=new SumPreviewForm(plan,data,(r,c)=>ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(()=>{
                    try{CenterOn((object)sheet,r,c);}catch(Exception ex){MessageBox.Show(ex.Message,"Excel 定位");}
                }));
                activePreview=preview;
                object workbook=(object)sheet.Parent;
                bool submitting=false,usesOutline=settings.UseOutline;
                string outlineBefore=PreviewPolicy.OutlineSignature(data);
                int summaryBefore=(int)sheet.Outline.SummaryRow;
                preview.FormClosed+=(s,e)=>{if(Object.ReferenceEquals(activePreview,preview))activePreview=null;preview.Dispose();};
                preview.SubmitRequested+=()=>{
                    if(submitting)return;submitting=true;preview.Enabled=false;
                    var selected=preview.Selected.ToList();
                    ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(()=>{
                        try{if(preview.IsDisposed)return;if(Commit((object)sheet,workbook,data,plan,selected,before,outlineBefore,summaryBefore,usesOutline))preview.Close();}
                        catch(Exception ex){MessageBox.Show(ex.Message,"智能求和 — 未完成",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
                        finally{submitting=false;if(!preview.IsDisposed)preview.Enabled=true;}
                    });
                };
                preview.Show(new ExcelOwner());
            }
        }
        private static SumPreviewForm activePreview;
        internal static void ClosePreview(){if(activePreview!=null&&!activePreview.IsDisposed)activePreview.Close();activePreview=null;}
        private static bool Commit(object sheetObject,object workbook,SheetData data,SumPlan plan,List<SumProposal> selected,string before,string outlineBefore,int summaryBefore,bool usesOutline)
        {
            dynamic app=ExcelBridge.App;dynamic sheet=sheetObject;
            if(!Object.ReferenceEquals((object)app.ActiveWorkbook,workbook))throw new InvalidOperationException("请切回生成预览时的工作簿后提交。没有写入。");
            if((bool)sheet.ProtectContents||(bool)sheet.FilterMode)throw new InvalidOperationException("请取消工作表保护并清除筛选后重新生成预览。");
            var current=ExcelBridge.Read(sheetObject);
            if(SnapshotGuard.Fingerprint(current)!=before || usesOutline&&(PreviewPolicy.OutlineSignature(current)!=outlineBefore||(int)sheet.Outline.SummaryRow!=summaryBefore))
                throw new InvalidOperationException("预览后源数据、公式、范围或分级发生变化，请关闭本窗口并重新生成预览。没有写入。");
                    var edits=new List<FormulaEdit>();
                    foreach(var proposal in plan.Proposals)
                    {
                        object value=At(data,proposal.Row,proposal.Col);string original=null;Dictionary<int,string> fr;
                        if(data.Formulas.TryGetValue(proposal.Row,out fr)&&fr.ContainsKey(proposal.Col))original="="+fr[proposal.Col];
                        var edit=new FormulaEdit{Row=proposal.Row,Col=proposal.Col,OriginalValue=value,OriginalFormula=original,NewFormula=proposal.Formula};
                        foreach(int childRow in proposal.ChildRows)
                        {
                            object cv=At(data,childRow,proposal.Col);
                            Dictionary<int,string> cf;bool hasFormula=data.Formulas.TryGetValue(childRow,out cf)&&cf.ContainsKey(proposal.Col);
                            if(!hasFormula && IsBlank(cv) && plan.Proposals.Any(p=>p.Row==childRow&&p.Col==proposal.Col))edit.RequiredAddresses.Add(SumSettingsForm.Col(proposal.Col)+childRow);
                        }
                        edits.Add(edit);
                    }
                    var addresses=new HashSet<string>(selected.Select(p=>p.Address),StringComparer.Ordinal);
                    FormulaTransaction.ValidateSelection(edits,addresses);
                    var chosen=edits.Where(e=>addresses.Contains(e.Address)).OrderByDescending(e=>e.Row).ThenBy(e=>e.Col).ToList();
                    foreach(var edit in chosen)CheckCell((object)sheet,edit);
                    if(MessageBox.Show("将写入 "+chosen.Count+" 个公式，其中原有数值/公式将按预览替换。\r\n不修改未勾选单元格，不自动保存。建议先保存副本。\r\nExcel 的 Ctrl+Z 可能被清空；可用“撤销上次求和”。","最终确认",MessageBoxButtons.OKCancel,MessageBoxIcon.Warning)!=DialogResult.OK)return false;
                    bool screen=app.ScreenUpdating,events=app.EnableEvents;int calculation=(int)app.Calculation;
                    try
                    {
                        app.ScreenUpdating=false;app.EnableEvents=false;app.Calculation=-4135; // xlCalculationManual
                        FormulaTransaction.Apply(chosen,e=>Write((object)sheet,e),e=>RestoreCell((object)sheet,e));
                        undoSheet=(object)sheet;undoWorkbook=(object)sheet.Parent;undo=chosen;
                        undoFingerprint=SnapshotGuard.Fingerprint(ExcelBridge.Read((object)sheet));
                    }
                    finally{app.Calculation=calculation;app.EnableEvents=events;app.ScreenUpdating=screen;}
                    // Recalculation may change cached Value2; capture only after calculation mode is restored.
                    undoFingerprint=SnapshotGuard.Fingerprint(ExcelBridge.Read((object)sheet));
            MessageBox.Show("已写入 "+chosen.Count+" 个公式。工作簿没有自动保存。","智能求和完成");return true;
        }
        private static void CenterOn(object sheetObject,int row,int col)
        {
            dynamic app=ExcelBridge.App;dynamic sheet=sheetObject;bool events=app.EnableEvents;
            try
            {
                app.EnableEvents=false;sheet.Parent.Activate();sheet.Activate();app.Goto(sheet.Cells[row,col],false);
                dynamic window=app.ActiveWindow;
                int fr=(bool)window.FreezePanes?(int)window.SplitRow:0,fc=(bool)window.FreezePanes?(int)window.SplitColumn:0;
                dynamic pane=window.Panes[window.Panes.Count];
                int rows=(int)pane.VisibleRange.Rows.Count,cols=(int)pane.VisibleRange.Columns.Count;
                if(row>fr)pane.ScrollRow=PreviewPolicy.CenterStart(row,rows,fr,1048576);
                if(col>fc)pane.ScrollColumn=PreviewPolicy.CenterStart(col,cols,fc,16384);
            }
            finally{app.EnableEvents=events;}
        }
        private static bool IsBlank(object value){return value==null || value is string && string.IsNullOrWhiteSpace((string)value);}
        private static object At(SheetData data,int row,int col){Dictionary<int,object> r;object v;return data.Values.TryGetValue(row,out r)&&r.TryGetValue(col,out v)?v:null;}
        private static void CheckCell(object sheetObject,FormulaEdit e)
        {
            dynamic sheet=sheetObject;dynamic cell=sheet.Cells[e.Row,e.Col];
            if((bool)cell.MergeCells)throw new InvalidOperationException(e.Address+" 是合并单元格，请先人工处理；本批未写入。");
            if((bool)cell.HasArray)throw new InvalidOperationException(e.Address+" 属于数组公式，请先人工处理；本批未写入。");
            // HasSpill is unavailable on older Excel; actual write still has transaction rollback protection.
            try{if((bool)cell.HasSpill)throw new InvalidOperationException(e.Address+" 属于溢出公式，请先人工处理；本批未写入。");}
            catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}
            catch(System.Runtime.InteropServices.COMException ex){if(ex.ErrorCode!=unchecked((int)0x80020003)&&ex.ErrorCode!=unchecked((int)0x80020006))throw;}
            foreach(dynamic table in sheet.ListObjects)
            {
                dynamic intersection=ExcelBridge.App.Intersect(cell,table.Range);
                if(intersection!=null)throw new InvalidOperationException(e.Address+" 位于 Excel 表对象中，写入可能自动扩展整列；本批已停止，请先转换为普通区域。");
            }
        }
        private static void Write(object sheetObject,FormulaEdit e)
        {
            dynamic sheet=sheetObject;dynamic cell=sheet.Cells[e.Row,e.Col];cell.Formula=e.NewFormula;
            if(!(bool)cell.HasFormula || !string.Equals(Convert.ToString(cell.Formula),e.NewFormula,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException(e.Address+" 公式写入回读不一致（可能为文本格式）。");
        }
        private static void RestoreCell(object sheetObject,FormulaEdit e)
        {
            dynamic sheet=sheetObject;dynamic cell=sheet.Cells[e.Row,e.Col];
            if(e.OriginalFormula!=null)cell.Formula=e.OriginalFormula;
            else if(e.OriginalValue==null)cell.ClearContents();
            else cell.Value2=e.OriginalValue;
            if(e.OriginalFormula!=null&&!string.Equals(Convert.ToString(cell.Formula),e.OriginalFormula,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException(e.Address+" 原公式恢复回读不一致。");
            if(e.OriginalFormula==null&&!Object.Equals((object)cell.Value2,e.OriginalValue))throw new InvalidOperationException(e.Address+" 原值恢复回读不一致。");
        }
        internal static void Undo()
        {
            if(undo==null||undo.Count==0)throw new InvalidOperationException("当前会话没有可撤销的求和记录。");
            dynamic app=ExcelBridge.App;dynamic sheet=undoSheet;
            if(!Object.ReferenceEquals((object)app.ActiveWorkbook,undoWorkbook))throw new InvalidOperationException("请切回上次智能求和的工作簿。");
            if((bool)sheet.ProtectContents || (bool)sheet.FilterMode)throw new InvalidOperationException("请先取消工作表保护并清除筛选条件。");
            if(SnapshotGuard.Fingerprint(ExcelBridge.Read((object)sheet))!=undoFingerprint)throw new InvalidOperationException("求和后数据、公式或范围已变化；为避免覆盖后续编辑，已拒绝撤销。请从保存的副本恢复。");
            foreach(var e in undo)CheckCell((object)sheet,e);
            if(MessageBox.Show("恢复上次求和写入的 "+undo.Count+" 个单元格原内容？","撤销上次求和",MessageBoxButtons.OKCancel,MessageBoxIcon.Warning)!=DialogResult.OK)return;
            bool screen=app.ScreenUpdating,events=app.EnableEvents;int calc=(int)app.Calculation;
            var restoreEdits=undo.AsEnumerable().Reverse().Select(e=>new FormulaEdit {Row=e.Row,Col=e.Col,
                OriginalFormula=e.NewFormula,OriginalValue=null,NewFormula=e.OriginalFormula??"[restore original value]"}).ToList();
            var originals=undo.ToDictionary(e=>e.Address,StringComparer.Ordinal);
            try
            {
                app.ScreenUpdating=false;app.EnableEvents=false;app.Calculation=-4135;
                FormulaTransaction.Apply(restoreEdits,e=>RestoreCell((object)sheet,originals[e.Address]),e=>Write((object)sheet,originals[e.Address]));
                undo=null;undoSheet=null;undoWorkbook=null;undoFingerprint=null;
            }
            finally{app.Calculation=calc;app.EnableEvents=events;app.ScreenUpdating=screen;}
        }
    }
}
