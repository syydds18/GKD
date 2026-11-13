using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Outline.Core;

namespace Outline.Addin
{
    internal static class PreviewPolicy
    {
        internal static bool IsBlank(SumProposal p){return !p.HasExistingFormula&&string.IsNullOrWhiteSpace(p.OriginalText);}
        internal static string Kind(SumProposal p){return p.HasExistingFormula?"已有公式":IsBlank(p)?"原值为空":"已有数值";}
        internal static IEnumerable<int> BlankIndices(IList<SumProposal> items){return Enumerable.Range(0,items.Count).Where(i=>IsBlank(items[i]));}
        internal static bool[] Selection(IList<SumProposal> items,string mode)
        {
            if(mode!="all"&&mode!="none"&&mode!="blank")throw new ArgumentException("Unknown selection mode",nameof(mode));
            return items.Select(p=>mode=="all"||mode=="blank"&&IsBlank(p)).ToArray();
        }
        internal static string OutlineSignature(SheetData data){return string.Join(";",data.OutlineLevels.OrderBy(p=>p.Key).Select(p=>p.Key+":"+p.Value));}
        internal static int CenterStart(int target,int visible,int frozen,int maximum)
        {return Math.Max(frozen+1,Math.Min(maximum-Math.Max(1,visible)+1,target-Math.Max(1,visible)/2));}
        internal static int NextExisting(IList<SumProposal> items,int current)
        {
            for(int step=1;step<=items.Count;step++){int i=(Math.Max(-1,current)+step)%items.Count;if(!IsBlank(items[i]))return i;}
            return -1;
        }
        internal static string Col(int c){string s="";while(c>0){c--;s=(char)('A'+c%26)+s;c/=26;}return s;}
        internal static string RowPreview(SheetData data,int row)
        {
            var text=new StringBuilder("原工作表快照 · 第"+row+"行\r\n");int shown=0;
            Dictionary<int,object> values;Dictionary<int,string> formulas;
            data.Values.TryGetValue(row,out values);data.Formulas.TryGetValue(row,out formulas);
            var cols=new SortedSet<int>();if(values!=null)foreach(int c in values.Keys)cols.Add(c);if(formulas!=null)foreach(int c in formulas.Keys)cols.Add(c);
            foreach(int c in cols)
            {
                object value=null;string formula=null;if(values!=null)values.TryGetValue(c,out value);if(formulas!=null)formulas.TryGetValue(c,out formula);
                if(value==null&&string.IsNullOrWhiteSpace(formula))continue;
                string part=Col(c)+row+"："+Convert.ToString(value,CultureInfo.InvariantCulture);
                if(!string.IsNullOrWhiteSpace(formula))part+="  ["+(formula.StartsWith("=",StringComparison.Ordinal)?formula:"="+formula)+"]";
                text.AppendLine(part.Length>220?part.Substring(0,220)+"…":part);
                if(++shown>=24){text.AppendLine("更多列请双击定位 Excel 查看。");break;}
            }
            return text.ToString();
        }
    }
}
