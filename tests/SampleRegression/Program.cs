using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Outline.Core;
var input=JsonDocument.Parse(System.IO.File.ReadAllText(args[0])).RootElement;
var sheet=new SheetData{MaxRow=input.GetProperty("maxRow").GetInt32(),MaxCol=input.GetProperty("maxCol").GetInt32()};
foreach(var row in input.GetProperty("rows").EnumerateArray()){
 int r=row.GetProperty("row").GetInt32();sheet.OutlineLevels[r]=row.GetProperty("level").GetInt32();
 sheet.Values[r]=new Dictionary<int,object>();foreach(var c in row.GetProperty("cells").EnumerateObject())sheet.Values[r][int.Parse(c.Name)]=c.Value.ValueKind==JsonValueKind.Number?(object)c.Value.GetDouble():c.Value.ToString();
}
var plan=SmartSumPlanner.Build(sheet,new Detection{HeaderRow=3,NameCol=3,CodeCol=0},new[]{7},true);
void Check(bool ok,string message){if(!ok)throw new Exception(message);}
string Formula(int row)=>plan.Proposals.Single(x=>x.Row==row).Formula;
Check(plan.Proposals.Count==255,"all 255 native parents must receive proposals");
Check(plan.Proposals.Count(p=>p.ChildRows.Count==1)==44,"all 44 single-child parents included");
Check(Formula(1387)=="=SUM(G1388,G1390)","gap region branches both included");
Check(Formula(1388)=="=SUM(G1389)","single-child branch included");
Check(Formula(1390)=="=SUM(G1391:G1393)","branch leaves included");
Check(Formula(1384)=="=SUM(G1385,G1387,G1394,G1396)","ancestor totals direct branches only");
Check(Formula(68)=="=SUM(G69)"&&Formula(1361)=="=SUM(G1362)"&&Formula(1827)=="=SUM(G1828)","reported single-child cases");
var detected=Engine.Detect(sheet);var compact=OutlineHierarchy.Compact(detected.Levels);
var stack=new List<int>();foreach(var pair in compact.OrderBy(x=>x.Key)){
 while(stack.Count>0&&stack[stack.Count-1]>=pair.Value)stack.RemoveAt(stack.Count-1);
 Check(stack.Count==pair.Value,"generated outline must not contain absent intermediate depth at row "+pair.Key);stack.Add(pair.Value);
}
Check(compact[1388]==2&&compact[1390]==2,"smart grouping pipeline must compact raw dotted depths");
var rows=plan.Proposals.ToDictionary(x=>x.Row);double Eval(int row){return rows.ContainsKey(row)?rows[row].ChildRows.Sum(Eval):Convert.ToDouble(sheet.Values[row].ContainsKey(7)?sheet.Values[row][7]:0);}
Check(Math.Abs(Eval(1387)-79.78473)<1e-8,"structural total not old J value");
Check(Math.Abs(Eval(1384)-1144.6054)<1e-8,"top total matches leaf total");
var report=new{count=plan.Proposals.Count,warningCount=plan.Warnings.Count,levels=new[]{1387,1388,1389,1390,1391}.Select(r=>new{row=r,raw=detected.Levels[r],compact=compact[r]}),proposals=plan.Proposals.Select(p=>new{p.Row,p.Address,p.Formula,p.ChildRows}),warnings=plan.Warnings};
System.IO.File.WriteAllText(args[1],JsonSerializer.Serialize(report,new JsonSerializerOptions{WriteIndented=true}));
Console.WriteLine("PASS: real 1840-row sample; 255 native parent proposals; 44 single-child cases included; smart-group outline gap removed; affected formulas and recursive amounts verified.");
