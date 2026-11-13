using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Outline.Core
{
    public sealed class SumPlan
    {
        public List<SumProposal> Proposals { get; set; } = new List<SumProposal>();
        public List<string> Warnings { get; set; } = new List<string>();
    }
    public sealed class SumProposal
    {
        public int Row { get; set; }
        public int Col { get; set; }
        public int Level { get; set; }
        public string Address { get; set; }
        public string RowName { get; set; }
        public string Formula { get; set; }
        public string OriginalText { get; set; }
        public List<int> ChildRows { get; set; } = new List<int>();
        public bool HasExistingFormula { get; set; }
    }

    /// <summary>Pure preview planner. Never changes SheetData or writes formulas.</summary>
    public static class SmartSumPlanner
    {
        private sealed class Node
        {
            public int Row;
            public int Level;
            public string Name;
            public bool Ambiguous;
            public CodeInfo Code;
            public HashSet<string> ChildCodes = new HashSet<string>(StringComparer.Ordinal);
            public List<Node> Children = new List<Node>();
        }
        public static Detection FindOutlineLayout(SheetData sheet)
        {
            if(sheet==null)throw new ArgumentNullException(nameof(sheet));
            string[] labels={"名称","项目名称","工程项目","费用名称"};
            for(int r=1;r<=Math.Min(30,sheet.MaxRow);r++)
            for(int c=1;c<=sheet.MaxCol;c++)
            {
                string text=Engine.Norm(Value(sheet,r,c)).Replace(" ","");
                if(labels.Contains(text))return new Detection {HeaderRow=r,NameCol=c,CodeCol=0};
            }
            return null;
        }
        public static SumPlan Build(SheetData sheet, Detection detection, IEnumerable<int> columns, bool useExistingOutline = false)
        {
            if(sheet==null) throw new ArgumentNullException(nameof(sheet));
            if(detection==null) throw new ArgumentNullException(nameof(detection));
            if(columns==null) throw new ArgumentNullException(nameof(columns));
            if(sheet.MaxRow<0 || sheet.MaxRow>1048576) throw new ArgumentOutOfRangeException(nameof(sheet),"Excel行范围为1..1048576。");
            if(sheet.MaxCol<0 || sheet.MaxCol>16384) throw new ArgumentOutOfRangeException(nameof(sheet),"Excel列范围为1..16384。");
            var plan = new SumPlan();
            var cols = columns.Distinct().OrderBy(c => c).ToList();
            if(cols.Any(c=>c<1 || c>16384)) throw new ArgumentOutOfRangeException(nameof(columns),"Excel列范围为1..16384。");
            if(cols.Count==0) { plan.Warnings.Add("未选择求和列。"); return plan; }
            if(detection.HeaderRow<0 || detection.HeaderRow>=sheet.MaxRow || detection.NameCol<1 || detection.NameCol>sheet.MaxCol ||
                !useExistingOutline && (detection.HeaderRow<1 || detection.CodeCol<1 || detection.CodeCol>16384 || detection.CodeCol==detection.NameCol))
            { plan.Warnings.Add(useExistingOutline?"请确认数据起始行和名称列，未生成提案。":"未识别到可靠表头、编号列和名称列，未生成提案。"); return plan; }
            // Caller preflights Excel SummaryRow == above when outline is selected.
            // Never substitute another source, even if the chosen source is flat/missing.
            var levels=useExistingOutline ? sheet.OutlineLevels : detection.Levels;
            if(useExistingOutline && levels!=null)
            {
                var compact=OutlineHierarchy.Compact(levels.Where(x=>x.Key>detection.HeaderRow).ToDictionary(x=>x.Key,x=>x.Value));
                foreach(var entry in compact)
                    if(levels[entry.Key]!=entry.Value)plan.Warnings.Add("第"+entry.Key+"行：原生层级空档已按实际嵌套压紧（"+levels[entry.Key]+" → "+entry.Value+"），仅用于本次预览，不修改原分级；请确认父子归属。");
                levels=compact;
            }
            bool hasPartBanner=!useExistingOutline && Enumerable.Range(detection.HeaderRow+1,sheet.MaxRow-detection.HeaderRow).Any(r=>
                Engine.CodeFamily(Value(sheet,r,detection.CodeCol)).Family=="part" || Blank(Value(sheet,r,detection.CodeCol)) && Engine.CodeFamily(Value(sheet,r,detection.NameCol)).Family=="part");
            if(levels==null || !levels.Values.Any(v=>v>0) && !hasPartBanner)
            { plan.Warnings.Add("所选层级来源没有可靠父子层级，未生成求和提案。"); return plan; }
            var nodes = new List<Node>(); var stack = new List<Node>();
            bool bannerActive=false; int? bannerOffset=null;
            for (int r = detection.HeaderRow + 1; r <= sheet.MaxRow; r++)
            {
                var code=Engine.CodeFamily(Value(sheet,r,detection.CodeCol));
                string name=Engine.Norm(Value(sheet,r,detection.NameCol));
                bool namedPart=Blank(Value(sheet,r,detection.CodeCol)) && Engine.CodeFamily(name).Family=="part";
                bool part=code.Family=="part" || namedPart;
                if ((!useExistingOutline && code.Family == null && !namedPart) || name.Length == 0)
                {
                    Dictionary<int,object> valueRow; Dictionary<int,string> formulaRow;
                    bool occupied=sheet.Values!=null && sheet.Values.TryGetValue(r,out valueRow) && valueRow!=null && valueRow.Values.Any(v=>!Blank(v)) ||
                        sheet.Formulas!=null && sheet.Formulas.TryGetValue(r,out formulaRow) && formulaRow!=null && formulaRow.Values.Any(f=>!string.IsNullOrWhiteSpace(f));
                    if (occupied)
                    {
                        foreach(var ancestor in stack) ancestor.Ambiguous=true;
                        plan.Warnings.Add("第"+r+"行："+(useExistingOutline?"名称缺失":"编号缺失/未知编号或名称缺失（小计/费用/统计等歧义行）")+"；跳过相关父项。");
                    }
                    continue;
                }
                int level;
                if (!levels.TryGetValue(r,out level) || level<0 || level>7)
                {
                    foreach(var ancestor in stack) ancestor.Ambiguous=true;
                    plan.Warnings.Add("第"+r+"行：层级缺失或超出0..7，归属歧义，跳过相关父项。");
                    continue;
                }
                // Detection represents standalone/merged part banners at level 0 and
                // may also start the first numbered family at 0. Add only the explicit
                // part edge; preserve all relative detected family/dotted boundaries.
                if(!useExistingOutline)
                {
                    if(part) { bannerActive=namedPart || !detection.PartIsFamily; bannerOffset=null; }
                    else if(bannerActive)
                    {
                        if(!bannerOffset.HasValue) bannerOffset=level==0 ? 1 : 0;
                        level+=bannerOffset.Value;
                        if(level>7)
                        {
                            foreach(var ancestor in stack) ancestor.Ambiguous=true;
                            plan.Warnings.Add("第"+r+"行：部分标题下层级超过7，跳过歧义范围。"); continue;
                        }
                    }
                }
                var n = new Node { Row=r, Level=level, Name=name, Code=namedPart ? Engine.CodeFamily(name) : code };
                while (stack.Count > 0 && stack[stack.Count-1].Level >= level) stack.RemoveAt(stack.Count-1);
                if(!useExistingOutline && stack.Count>0 && code.Family=="dotted" && !DottedFits(stack[stack.Count-1].Code,code))
                {
                    foreach(var ancestor in stack) ancestor.Ambiguous=true;
                    plan.Warnings.Add("第"+r+"行：点号编号前缀与直接父项不匹配，归属歧义，跳过相关父项。");
                    n.Ambiguous=true;
                }
                if(stack.Count>0)
                {
                    var parent=stack[stack.Count-1];
                    if(parent.Ambiguous) n.Ambiguous=true;
                    bool duplicate=!useExistingOutline && !parent.ChildCodes.Add(n.Code.Family+":"+n.Code.Text);
                    if(duplicate || level!=parent.Level+1)
                    {
                        foreach(var ancestor in stack) ancestor.Ambiguous=true;
                        plan.Warnings.Add("第"+r+"行："+(duplicate?"重复编号":"层级跳跃缺失中间节点")+"，归属歧义，跳过相关父项。");
                    }
                    parent.Children.Add(n);
                }
                stack.Add(n); nodes.Add(n);
            }
            foreach (var n in nodes.AsEnumerable().Reverse())
            foreach (int c in cols)
            {
                if(n.Ambiguous) continue;
                if (n.Children.Count == 0) continue;
                object target=Value(sheet,n.Row,c);
                if(!Blank(target) && !Engine.ToNumberValue(target).Ok)
                { plan.Warnings.Add(Engine.ColLetter(c)+n.Row+"：父项为错误值/文字/布尔值，跳过。"); continue; }
                bool safe=true;
                foreach(var child in n.Children)
                {
                    object raw=Value(sheet,child.Row,c);
                    if(!Blank(raw) && (raw is string || !Engine.ToNumberValue(raw).Ok))
                    { plan.Warnings.Add(Engine.ColLetter(c)+child.Row+"：子项为数字文本/错误值/文字/布尔值，SUM可能忽略，跳过父项。"); safe=false; }
                    if(Blank(raw) && ExistingFormula(sheet,child.Row,c).Length==0 && child.Children.Count>0 && !plan.Proposals.Any(p=>p.Row==child.Row && p.Col==c))
                    { plan.Warnings.Add(Engine.ColLetter(c)+n.Row+"：空白子分组缺少有效求和提案，跳过上级以免漏计。"); safe=false; }
                }
                if(!safe) continue;
                var rows = n.Children.Select(x=>x.Row).Distinct().OrderBy(r=>r).ToList();
                int segments=1;
                for(int i=1;i<rows.Count;i++) if(rows[i]!=rows[i-1]+1) segments++;
                if(segments>255)
                { plan.Warnings.Add(Engine.ColLetter(c)+n.Row+"：SUM参数段超过255，保守跳过。"); continue; }
                if(plan.Proposals.Count==5000) throw new InvalidOperationException("求和候选超过5000，已拒绝整个计划；请缩小范围或减少列。");
                plan.Proposals.Add(new SumProposal { Row=n.Row, Col=c, Level=n.Level, Address=Engine.ColLetter(c)+n.Row,
                    RowName=n.Name, ChildRows=rows, Formula=Formula(c,rows), HasExistingFormula=ExistingFormula(sheet,n.Row,c).Length>0,
                    OriginalText=ExistingFormula(sheet,n.Row,c).Length>0 ? ExistingFormula(sheet,n.Row,c) : target as string ?? Engine.Norm(target) });
            }
            return plan;
        }
        private static bool DottedFits(CodeInfo parent,CodeInfo child)
        {
            if(parent.Family=="dotted")
                return child.Parts.Count==parent.Parts.Count+1 && parent.Parts.SequenceEqual(child.Parts.Take(parent.Parts.Count));
            if(parent.Family=="num")
            {
                string digits=new string(parent.Text.TakeWhile(ch=>ch>='0' && ch<='9').ToArray());
                int ordinal;
                return int.TryParse(digits,NumberStyles.None,CultureInfo.InvariantCulture,out ordinal) && child.Parts.Count==2 && child.Parts[0]==ordinal;
            }
            // No explicit numeric base: do not invent a missing dotted parent.
            return child.Parts.Count==2;
        }
        private static bool Blank(object value) { return value == null || value is string && Engine.Norm(value).Length == 0; }
        private static string ExistingFormula(SheetData sheet,int row,int col)
        {
            Dictionary<int,string> cells; string formula;
            return sheet.Formulas!=null && sheet.Formulas.TryGetValue(row,out cells) && cells!=null && cells.TryGetValue(col,out formula) ? formula ?? "" : "";
        }
        private static object Value(SheetData sheet, int row, int col)
        {
            Dictionary<int,object> cells; object value;
            return sheet.Values != null && sheet.Values.TryGetValue(row,out cells) && cells != null && cells.TryGetValue(col,out value) ? value : null;
        }
        private static string Formula(int col, List<int> rows)
        {
            string letter=Engine.ColLetter(col); var parts=new List<string>();
            for(int i=0;i<rows.Count;i++)
            {
                int start=rows[i], end=start;
                while(i+1<rows.Count && rows[i+1]==end+1) end=rows[++i];
                parts.Add(letter+start.ToString(CultureInfo.InvariantCulture)+(end>start?":"+letter+end.ToString(CultureInfo.InvariantCulture):""));
            }
            // Each argument is at most XFD1048576:XFD1048576 (21 chars).
            // With <=255 arguments the formula is at most 5615 chars (<8192).
            return "=SUM("+string.Join(",",parts)+")";
        }
    }
}
