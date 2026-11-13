using System;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using Outline.Core;

static class Program
{
    static readonly Dictionary<string, Action> Tests = new Dictionary<string, Action>();
    static int Main(string[] args)
    {
        Tests.Add("outline_layout_without_number_header", () => {
            var s=Sheet();s.Values.Clear();s.MaxRow=4;s.MaxCol=2;
            s.Values[1]=new Dictionary<int,object>{{1,"项目名称"},{2,"金额"}};
            s.Values[2]=new Dictionary<int,object>{{1,"合计"}};
            s.Values[3]=new Dictionary<int,object>{{1,"明细甲"},{2,10}};
            s.Values[4]=new Dictionary<int,object>{{1,"明细乙"},{2,20}};
            s.OutlineLevels=new Dictionary<int,int>{{1,0},{2,0},{3,1},{4,1}};
            var d=SmartSumPlanner.FindOutlineLayout(s);
            Check(d.HeaderRow==1&&d.NameCol==1&&d.CodeCol==0,"name header must not depend on number header");
            Check(Get<string>(At(Plan(s,new[]{2},true,d),2,2),"Formula")=="=SUM(B3:B4)","two-column sheet without number column");
            s.Values[1][1]="自定义标签";
            Check(SmartSumPlanner.FindOutlineLayout(s)==null,"unknown header must require manual start/name selection");
            s.Values.Remove(1);s.Values[1]=s.Values[2];s.Values[2]=s.Values[3];s.Values[3]=s.Values[4];s.Values.Remove(4);s.MaxRow=3;
            s.OutlineLevels=new Dictionary<int,int>{{1,0},{2,1},{3,1}};
            d=new Detection{HeaderRow=0,NameCol=1,CodeCol=0};
            Check(Get<string>(At(Plan(s,new[]{2},true,d),1,2),"Formula")=="=SUM(B2:B3)","manual start at row1 must support headerless outline");
        });
        Tests.Add("outline_blank_codes_and_no_code_column", () => {
            var s=Tree();foreach(var row in s.Values.Where(x=>x.Key>1))row.Value[1]=null;
            s.OutlineLevels=new Dictionary<int,int>{{2,0},{3,1},{4,2},{5,2},{6,1}};
            var d=new Detection{HeaderRow=1,CodeCol=0,NameCol=2};
            var p=Plan(s,outline:true,d:d);
            Check(Get<string>(At(p,2),"Formula")=="=SUM(C3,C6)","outline must allow entirely blank codes and absent code column");
            Check(Get<string>(At(p,3),"Formula")=="=SUM(C4:C5)","outline direct children without codes");
            Check(Proposals(Plan(s,outline:false,d:d)).Count==0,"number mode must still require code column");
            foreach(var row in s.Values.Where(x=>x.Key>1))row.Value[1]="9.9";
            Check(Proposals(Plan(s,outline:true,d:d)).Count==2,"outline must ignore duplicate and conflicting codes");
            s.Values[4][2]=null;
            Check(Proposals(Plan(s,outline:true,d:d)).Count==0,"missing name remains a safety barrier");
        });
        Tests.Add("outline_gap_compacts_actual_parent_tree", () => {
            var s=Sheet(new object[]{"第四项","总",null},new object[]{1,"水系统",10},new object[]{2,"接地",null},
                new object[]{"2.2.1","机电消防",null},new object[]{null,"管路",35.906895},
                new object[]{"2.2.2","建筑消防",null},new object[]{null,"管甲",9.827835},new object[]{null,"管乙",15.6},new object[]{null,"管丙",18.45});
            s.OutlineLevels=new Dictionary<int,int>{{2,0},{3,1},{4,1},{5,3},{6,4},{7,3},{8,4},{9,4},{10,4}};
            var compact=OutlineHierarchy.Compact(s.OutlineLevels);
            Check(compact[5]==2&&compact[6]==3&&compact[7]==2&&compact[8]==3,"gap removed consistently across sibling branches");
            Check(s.OutlineLevels[5]==3,"normalization must not mutate source outline");
            Check(OutlineHierarchy.Compact(compact).OrderBy(x=>x.Key).SequenceEqual(compact.OrderBy(x=>x.Key)),"normalization idempotent");
            var p=Plan(s,outline:true);
            Check(Get<string>(At(p,4),"Formula")=="=SUM(C5,C7)","nearest enclosing parent references direct existing branches only");
            Check(Get<string>(At(p,5),"Formula")=="=SUM(C6)","single-child branch after gap");
            Check(Get<string>(At(p,7),"Formula")=="=SUM(C8:C10)","second orphan branch retained");
            Check(Get<string>(At(p,2),"Formula")=="=SUM(C3:C4)","upper parent no longer blocked");
            Check(Warnings(p).Any(x=>x.Contains("压紧")),"must disclose normalization instead of silently claiming numbering is complete");
        });
        Tests.Add("direct_children_three_levels", DirectChildren);
        Tests.Add("multiple_columns_existing_formula", MultipleColumns);
        Tests.Add("single_child_formula", SingleChild);
        Tests.Add("ambiguous_rows_block_ancestors", AmbiguousRows);
        Tests.Add("typed_amount_safety", () => {
            foreach(object bad in new object[]{"12","#VALUE!","#DIV/0!",true,"文字",double.NaN,double.PositiveInfinity,new System.Runtime.InteropServices.ErrorWrapper(2015)})
            {
                var s=Tree(); s.Values[4][3]=bad;
                var p=Plan(s,new[]{3,4});
                Check(!Proposals(p).Any(x=>Get<int>(x,"Col")==3),"unsafe child must block blank ancestors: "+bad);
                Check(Proposals(p).Count==2,"bad C must not block D"); Check(Warnings(p).Count>0,"unsafe warning");
            }
            foreach(object bad in new object[]{"#REF!","bad",false})
            {
                var s=Tree();s.Values[2][3]=bad;var p=Plan(s);
                Check(Proposals(p).Count==1 && Get<int>(Proposals(p)[0],"Row")==3,"bad parent target must skip");
            }
            var good=Tree();good.Values[2][3]="1,234";
            Check(Proposals(Plan(good)).Count==2,"parseable numeric parent is safe to replace");
        });
        Tests.Add("explicit_outline_source", () => {
            var s=Sheet(new object[]{"一","父",null},new object[]{"1","甲",10},new object[]{"2","乙",20});
            s.OutlineLevels=new Dictionary<int,int>{{2,0},{3,0},{4,0}};
            var p=Plan(s,outline:true);
            Check(Proposals(p).Count==0 && Warnings(p).Count>0,"explicit flat outline must not fall back to detected");
            s.OutlineLevels[3]=1;s.OutlineLevels[4]=1;
            var d=Engine.Detect(s);d.Levels=new Dictionary<int,int>{{2,0},{3,0},{4,0}};d.MaxLevel=0;
            Check(Proposals(Plan(s,outline:false,d:d)).Count==0,"default must not silently use outline");
            Check(Get<string>(At(Plan(s,outline:true,d:d),2),"Formula")=="=SUM(C3:C4)","explicit outline tree");
            s=Sheet(new object[]{"一","父",null},new object[]{"1","甲",10},new object[]{null,"小计",10},new object[]{"2","乙",20});
            s.OutlineLevels=new Dictionary<int,int>{{2,0},{3,1},{4,0},{5,1}};
            var split=Plan(s,outline:true);
            Check(Get<string>(At(split,2),"Formula")=="=SUM(C3)","outline subtotal must not attach following sibling to previous root");
            Check(Get<string>(At(split,4),"Formula")=="=SUM(C5)","single child after subtotal boundary belongs only to subtotal");
        });
        Tests.Add("input_and_excel_bounds", () => {
            foreach(int col in new[]{0,-1,16385})
                Throws<ArgumentOutOfRangeException>(()=>Plan(Tree(),new[]{col}),"illegal column "+col);
            var s=Tree();s.MaxRow=1048577;
            Throws<ArgumentOutOfRangeException>(()=>Plan(s),"Excel row bound");
            s=Tree();s.MaxCol=16385;
            Throws<ArgumentOutOfRangeException>(()=>Plan(s),"Excel column bound");
            s=Tree();s.MaxCol=16384;
            var p=Plan(s,new[]{16384});
            Check(Get<string>(At(p,2,16384),"Formula")=="=SUM(XFD3,XFD6)","last Excel column");
            Check(Warnings(Plan(Tree(),new int[0])).Count>0,"no columns warning");
            Throws<ArgumentNullException>(()=>Plan(null),"null sheet");
        });
        Tests.Add("missing_or_invalid_levels_are_barriers", () => {
            foreach(int invalid in new[]{-1,8,99})
            {
                var s=Tree();var d=Engine.Detect(s);d.Levels[4]=invalid;
                var p=Plan(s,d:d);Check(Proposals(p).Count==0 && Warnings(p).Count>0,"invalid level must not bridge");
            }
            {
                var s=Tree();var d=Engine.Detect(s);d.Levels.Remove(4);
                var p=Plan(s,d:d);Check(Proposals(p).Count==0 && Warnings(p).Count>0,"missing level must not bridge");
            }
            {
                var s=Tree();var d=Engine.Detect(s);d.HeaderRow=0;
                Check(Proposals(Plan(s,d:d)).Count==0,"missing header must refuse");
            }
            var uncoded=Sheet(new object[]{null,"合计",100},new object[]{null,"甲",10},new object[]{null,"乙",20});
            Check(Proposals(Plan(uncoded)).Count==0 && Warnings(Plan(uncoded)).Count>0,"uncoded low confidence");
        });
        Tests.Add("part_banner_is_structural", () => {
            var s=Sheet(new object[]{null,"第一部分 建筑",null},new object[]{"一","甲",10},new object[]{"二","乙",20});
            Check(Get<string>(At(Plan(s),2),"Formula")=="=SUM(C3:C4)","part banner in name must act as explicit part code");
        });
        Tests.Add("sum_argument_limit_255", () => {
            foreach(int count in new[]{255,256})
            {
                var s=Sheet();s.MaxRow=2+count*2;
                s.Values[2]=new Dictionary<int,object>{{1,"一"},{2,"父"}};
                var d=new Detection{HeaderRow=1,CodeCol=1,NameCol=2,MaxLevel=1};d.Levels[2]=0;
                for(int i=0;i<count;i++)
                {
                    int row=3+i*2; s.Values[row]=new Dictionary<int,object>{{1,(i+1).ToString()},{2,"叶"},{3,1}};d.Levels[row]=1;
                }
                var p=Plan(s,d:d);
                if(count==255) Check(Proposals(p).Count==1 && Get<string>(At(p,2),"Formula").Split(',').Length==255,"255 segments allowed");
                else Check(Proposals(p).Count==0 && Warnings(p).Any(x=>x.Contains("255")),"256 segments must skip with warning");
            }
        });
        Tests.Add("candidate_cap_is_atomic", () => {
            var s=Tree();s.MaxCol=2500;
            Check(Proposals(Plan(s,Enumerable.Range(5,2500).ToArray())).Count==5000,"5000 proposals allowed");
            Throws<InvalidOperationException>(()=>Plan(s,Enumerable.Range(5,2501).ToArray()),"candidate overflow must fail whole plan, not truncate");
        });
        Tests.Add("dotted_prefix_boundary", () => {
            var s=Sheet(new object[]{"一","总",null},new object[]{"1","组一",null},new object[]{"1.1","甲",10},new object[]{"1.2","乙",20},
                new object[]{"2","组二",null},new object[]{"1.3","错误归属",30},new object[]{"2.1","丙",40},new object[]{"2.2","丁",50});
            var p=Plan(s);
            Check(!Proposals(p).Any(x=>Get<int>(x,"Row")==6 || Get<int>(x,"Row")==2),"must not attach 1.3 to group 2");
            Check(Get<string>(At(p,3),"Formula")=="=SUM(C4:C5)","valid earlier group unaffected");
            Check(Warnings(p).Any(x=>x.Contains("点号")),"dotted prefix mismatch warning");
        });
        Tests.Add("duplicate_sibling_codes_are_ambiguous", () => {
            var s=Sheet(new object[]{"一","父",null},new object[]{"1","甲",10},new object[]{"1","乙",20});
            var p=Plan(s);Check(Proposals(p).Count==0 && Warnings(p).Any(x=>x.Contains("重复编号")),"duplicate numbered siblings must not be guessed");
        });
        Tests.Add("hierarchy_gaps_are_ambiguous", () => {
            var s=Tree();var d=Engine.Detect(s);d.Levels[3]=2;d.Levels[4]=3;d.Levels[5]=3;d.Levels[6]=2;
            var p=Plan(s,d:d);Check(!Proposals(p).Any(x=>Get<int>(x,"Row")==2) && Warnings(p).Count>0,"missing intermediate level must block root");
        });
        Tests.Add("ambiguous_scope_does_not_spawn_new_groups", () => {
            var s=Sheet(new object[]{"一","父",null},new object[]{"1","甲",null},new object[]{null,"小计",20},
                new object[]{"(1)","不明组",null},new object[]{"①","甲",10},new object[]{"②","乙",10});
            Check(Proposals(Plan(s)).Count==0,"no new group proposals inside ambiguous range");
        });
        Tests.Add("nonempty_unselected_cells_are_not_blank_rows", () => {
            var s=Sheet(new object[]{"一","父",null},new object[]{"1","甲",10},new object[]{null,null,null,"备注"},new object[]{"2","乙",20});
            Check(Proposals(Plan(s)).Count==0,"nonempty note-only row is not a blank gap");
        });
        Tests.Add("blank_leaf_and_group_dependencies", () => {
            var s=Tree();s.Values[4][3]=null;
            Check(Proposals(Plan(s)).Count==2,"blank leaf may be referenced");
            s.Values.Remove(5); // single-child intermediate must also receive a proposal, never count as zero
            var p=Plan(s);
            Check(Proposals(p).Count==2 && Get<string>(At(p,3),"Formula")=="=SUM(C4)","single-child intermediate must not block ancestors");
            s.Values[3][3]=10;
            Check(Proposals(Plan(s)).Count==2,"existing numeric intermediate may be offered for explicit rebuild and referenced directly");
            s=Sheet(new object[]{"一","父",null},new object[]{"1","甲",10},new object[]{null,null,null},new object[]{"2","乙",20});
            Check(Get<string>(At(Plan(s),2),"Formula")=="=SUM(C3,C5)","blank gap is ignored, never included in a range");
        });
        Tests.Add("bottom_up_order_and_formula_limits", () => {
            var p=Plan(Tree());
            Check(Get<int>(Proposals(p)[0],"Row")==3,"dependencies before parents");
            foreach(var item in Proposals(p))
            {
                Check(Get<string>(item,"Formula").Length<=8192,"formula size");
                Check(Get<List<int>>(item,"ChildRows").All(r=>r>Get<int>(item,"Row")),"only downward references");
            }
        });
        Tests.Add("preview_preserves_original_text", () => {
            var s=Tree();s.Values[2][3]=" 1,234 ";
            Check(Get<string>(At(Plan(s),2),"OriginalText")==" 1,234 ","preview must retain original cell text exactly");
        });
        Tests.Add("maximal_range_formula_stays_below_8192", () => {
            int header=1047810;
            var s=new SheetData {MaxRow=1048576,MaxCol=16384};
            var d=new Detection {HeaderRow=header,CodeCol=1,NameCol=2,MaxLevel=1};
            int parent=header+1;
            s.Values[parent]=new Dictionary<int,object>{{1,"一"},{2,"父"}};d.Levels[parent]=0;
            for(int i=0;i<255;i++) for(int j=0;j<2;j++)
            {
                int r=parent+1+i*3+j;
                s.Values[r]=new Dictionary<int,object>{{1,(i*2+j+1).ToString()},{2,"叶"}}; d.Levels[r]=1;
            }
            var formula=Get<string>(At(Plan(s,new[]{16384},d:d),parent,16384),"Formula");
            Check(formula.Length==5615 && formula.Length<=8192,"255 maximal range segments length");
            Check(formula.Split(',').Length==255,"argument boundary");
        });
        int failed = 0, ran = 0;
        foreach (var t in Tests)
        {
            if (args.Length > 0 && t.Key != args[0]) continue;
            ran++;
            try { t.Value(); Console.WriteLine("PASS " + t.Key); }
            catch (Exception ex) { failed++; Console.WriteLine("FAIL " + t.Key + ": " + ex.Message); }
        }
        Console.WriteLine("Tests: " + ran + ", passed: " + (ran - failed) + ", failed: " + failed);
        return failed == 0 && ran > 0 ? 0 : 1;
    }
    static void Throws<T>(Action action,string message) where T:Exception
    {
        try { action(); } catch(T) { return; }
        throw new Exception("Expected "+typeof(T).Name+": "+message);
    }
    static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
    static SheetData Sheet(params object[][] rows)
    {
        var s = new SheetData { MaxRow = rows.Length + 1, MaxCol = 4 };
        s.Values[1] = new Dictionary<int, object> { {1,"编号"}, {2,"名称"}, {3,"金额"}, {4,"第二金额"} };
        for (int i = 0; i < rows.Length; i++)
        {
            s.Values[i+2] = new Dictionary<int, object>();
            for (int c = 0; c < rows[i].Length; c++) s.Values[i+2][c+1] = rows[i][c];
        }
        return s;
    }
    static object Plan(SheetData s, int[] cols = null, bool outline = false, Detection d = null)
    {
        Type type = typeof(Engine).Assembly.GetType("Outline.Core.SmartSumPlanner");
        Check(type != null, "SmartSumPlanner API missing");
        try { return type.GetMethod("Build").Invoke(null, new object[] { s, d ?? Engine.Detect(s), cols ?? new[] {3}, outline }); }
        catch (TargetInvocationException ex) { throw ex.InnerException; }
    }
    static T Get<T>(object obj, string name) { return (T)obj.GetType().GetProperty(name).GetValue(obj); }
    static List<object> Proposals(object plan) { return ((System.Collections.IEnumerable)Get<object>(plan,"Proposals")).Cast<object>().ToList(); }
    static List<string> Warnings(object plan) { return Get<List<string>>(plan,"Warnings"); }
    static object At(object plan, int row, int col = 3) { return Proposals(plan).Single(x => Get<int>(x,"Row") == row && Get<int>(x,"Col") == col); }
    static SheetData Tree()
    {
        return Sheet(new object[]{"一","总项",null,100}, new object[]{"1","分组",null,50},
          new object[]{"1.1","叶甲",10,20}, new object[]{"1.2","叶乙",20,30}, new object[]{"2","叶丙",30,50});
    }
    static void AmbiguousRows()
    {
        foreach(string label in new[]{"小计","合计","总计","费用率","年度统计","subtotal","fee","summary","普通未编号行",""})
        foreach(string code in new[]{"","XYZ"})
        {
            var s=Sheet(new object[]{"一","父",30}, new object[]{"1","子甲",10},new object[]{code,label,10},new object[]{"2","子乙",20});
            var p=Plan(s);
            Check(Proposals(p).Count==0,"must not cross ambiguous row "+code+label);
            Check(Warnings(p).Any(x=>x.Contains("编号") && x.Contains("3") || x.Contains("歧义")),"ambiguity warning missing");
        }
    }
    static void SingleChild()
    {
        var p=Plan(Sheet(new object[]{"一","父",50},new object[]{"1","独子",50}));
        Check(Get<string>(At(p,2),"Formula")=="=SUM(C3)","single child must receive formula");
        var s=Sheet(new object[]{"一","父",null},new object[]{"1","独子父项",null},new object[]{"1.1","叶",50});
        foreach(bool native in new[]{false,true})
        {
            s.OutlineLevels=new Dictionary<int,int>{{2,0},{3,1},{4,2}};
            var chain=Plan(s,outline:native);
            Check(Get<string>(At(chain,3),"Formula")=="=SUM(C4)","single-child chain lower formula");
            Check(Get<string>(At(chain,2),"Formula")=="=SUM(C3)","single-child chain upper formula without double count");
        }
    }
    static void MultipleColumns()
    {
        var s=Tree(); s.Formulas[2]=new Dictionary<int,string>{{3,"=123"}};
        var p=Plan(s,new[]{4,3,3});
        Check(Proposals(p).Count==4,"dedupe columns, keep both amounts");
        Check(Get<bool>(At(p,2),"HasExistingFormula"),"existing formula should be offered for explicit preview selection");
        Check(Get<string>(At(p,2),"OriginalText")=="=123","preserve original formula text");
        Check(Get<string>(At(p,2,4),"OriginalText")=="100","preserve fixed parent amount");
        Check(Get<string>(At(p,2,4),"Formula")=="=SUM(D3,D6)","independent same-column formula");
    }
    static void DirectChildren()
    {
        var s = Tree(); var p = Plan(s);
        Check(Proposals(p).Count == 2, "expected two parent proposals");
        Check(Get<string>(At(p,2),"Formula") == "=SUM(C3,C6)", "root must not double-count grandchildren");
        Check(Get<string>(At(p,3),"Formula") == "=SUM(C4:C5)", "contiguous direct children must compress");
        Check(Get<List<int>>(At(p,2),"ChildRows").SequenceEqual(new[]{3,6}), "direct children metadata");
        Check(Get<string>(At(p,2),"Address") == "C2", "A1 address");
        Check(s.Formulas.Count == 0 && s.Values[2][3] == null, "planner mutated input");
    }
}
