using System;
using System.Collections.Generic;
using System.Linq;

namespace Outline.Addin
{
    internal sealed class FormulaEdit
    {
        public int Row, Col;
        public object OriginalValue;
        public string OriginalFormula, NewFormula;
        public List<string> RequiredAddresses=new List<string>();
        public string Address { get { return ColumnLetter(Col) + Row.ToString(); } }
        private static string ColumnLetter(int col)
        {
            if(col<1 || col>16384)throw new ArgumentOutOfRangeException("Col");
            var s="";while(col>0){col--;s=(char)('A'+col%26)+s;col/=26;}return s;
        }
    }
    internal static class FormulaTransaction
    {
        internal static void ValidateSelection(IEnumerable<FormulaEdit> all,IEnumerable<string> selection)
        {
            var selected=new HashSet<string>(selection,StringComparer.Ordinal);
            foreach(var edit in all)
                if(selected.Contains(edit.Address))foreach(var dependency in edit.RequiredAddresses)
                    if(!selected.Contains(dependency))throw new InvalidOperationException(edit.Address+" 依赖空白子汇总 "+dependency+"；请同时勾选该子项，避免漏计。");
        }
        internal static void Apply(IEnumerable<FormulaEdit> source,Action<FormulaEdit> write,Action<FormulaEdit> restore)
        {
            if(source==null||write==null||restore==null)throw new ArgumentNullException();
            var edits=source.ToList();var seen=new HashSet<string>(StringComparer.Ordinal);
            foreach(var e in edits)
            {
                if(e==null||e.Row<1||e.Row>1048576||e.Col<1||e.Col>16384||string.IsNullOrEmpty(e.NewFormula)||!seen.Add(e.Address))
                    throw new ArgumentException("智能求和写入计划包含非法或重复目标单元格。", "source");
            }
            var applied=new List<FormulaEdit>();
            try{foreach(var e in edits){write(e);applied.Add(e);}}
            catch(Exception writeError)
            {
                var rollbackErrors=new List<Exception>();
                for(int i=applied.Count-1;i>=0;i--)try{restore(applied[i]);}catch(Exception ex){rollbackErrors.Add(ex);}
                // The failing write may have changed its cell before throwing; restore it too.
                if(applied.Count<edits.Count)try{restore(edits[applied.Count]);}catch(Exception ex){rollbackErrors.Add(ex);}
                if(rollbackErrors.Count>0)throw new InvalidOperationException("公式写入失败，且部分单元格回退失败；请立即关闭工作簿且不要保存。",writeError);
                throw new InvalidOperationException("公式写入失败，已回退本批已写入单元格。",writeError);
            }
        }
    }
}
