using System;
using System.Collections.Generic;
using Outline.Addin;
int tests=0;
void Assert(bool x,string message){if(!x)throw new Exception(message);tests++;}
var state=new Dictionary<string,object>{{"C2",100.0},{"C3","=OLD()"}};
var edits=new List<FormulaEdit>{new FormulaEdit{Row=3,Col=3,OriginalFormula="=OLD()",NewFormula="=SUM(C4:C5)"},new FormulaEdit{Row=2,Col=3,OriginalValue=100.0,NewFormula="=SUM(C3,C6)"}};
FormulaTransaction.Apply(edits,e=>{state[e.Address]=e.NewFormula;},e=>{state[e.Address]=e.OriginalFormula??e.OriginalValue;});
Assert((string)state["C2"]=="=SUM(C3,C6)","successful write missing");
state["C2"]=100.0;state["C3"]="=OLD()";bool failed=false;
try{FormulaTransaction.Apply(edits,e=>{state[e.Address]=e.NewFormula;if(e.Row==2)throw new Exception("write failed after assignment");},e=>state[e.Address]=e.OriginalFormula??e.OriginalValue);}catch(InvalidOperationException e){failed=e.Message.Contains("回退");}
Assert(failed && (double)state["C2"]==100.0 && (string)state["C3"]=="=OLD()","failed write must restore all attempted cells");
var log=new List<string>();
try{FormulaTransaction.Apply(edits,e=>{if(e.Row==2)throw new Exception("write failed");},e=>{log.Add(e.Address);throw new Exception("rollback failed");});}catch(InvalidOperationException e){Assert(e.Message.Contains("回退失败"),"rollback failure hidden");}
Assert(log.Count==2,"must attempt all rollbacks even after one fails");
try{FormulaTransaction.Apply(new[]{edits[0],edits[0]},e=>throw new Exception("must not write"),e=>{});throw new Exception("duplicates accepted");}catch(ArgumentException){tests++;}
var parent=new FormulaEdit{Row=2,Col=3,NewFormula="=SUM(C3,C6)",RequiredAddresses=new List<string>{"C3"}};
var child=new FormulaEdit{Row=3,Col=3,NewFormula="=SUM(C4:C5)"};
try{FormulaTransaction.ValidateSelection(new[]{parent,child},new[]{"C2"});throw new Exception("blank parent dependency accepted");}catch(InvalidOperationException){tests++;}
FormulaTransaction.ValidateSelection(new[]{parent,child},new[]{"C2","C3"});tests++;
Console.WriteLine("PASS: "+tests+" formula transaction checks");
