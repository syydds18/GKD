using System;
using System.Threading.Tasks;
using Outline.Setup;

int prompts=0,requests=0,waits=0;
void Check(bool ok){if(!ok)throw new Exception("close gate failed");}
Check(await ExcelCloseGate.EnsureClosed(()=>0,n=>{prompts++;return true;},()=>requests++,()=>Task.CompletedTask,3));Check(prompts==0&&requests==0);
Check(!await ExcelCloseGate.EnsureClosed(()=>1,n=>{prompts++;return false;},()=>requests++,()=>Task.CompletedTask,3));Check(requests==0);
int running=1;
Check(await ExcelCloseGate.EnsureClosed(()=>running,n=>true,()=>{requests++;running=0;},()=>Task.CompletedTask,3));Check(requests==1);
Check(!await ExcelCloseGate.EnsureClosed(()=>1,n=>true,()=>requests++,()=>{waits++;return Task.CompletedTask;},3));Check(waits==3&&requests==2);
Check(await ExcelCloseGate.CloseNormally(()=>1,n=>false,()=>throw new Exception("must not close"),()=>Task.CompletedTask,1)==ExcelCloseResult.Cancelled);
Check(await ExcelCloseGate.CloseNormally(()=>1,n=>true,()=>{},()=>Task.CompletedTask,1)==ExcelCloseResult.StillRunning);
Check(await ExcelCloseGate.CloseNormally(()=>0,n=>throw new Exception("must not prompt"),()=>{},()=>Task.CompletedTask,1)==ExcelCloseResult.Closed);
requests=0;
Check(!await ExcelCloseGate.ForceClosed(()=>1,n=>false,()=>requests++,()=>Task.CompletedTask,1));Check(requests==0);
running=1;
Check(await ExcelCloseGate.ForceClosed(()=>running,n=>true,()=>{requests++;running=0;},()=>Task.CompletedTask,1));Check(requests==1);
Check(!await ExcelCloseGate.ForceClosed(()=>1,n=>true,()=>requests++,()=>Task.CompletedTask,1));Check(requests==2);
requests=0;
Check(await ExcelCloseGate.ForceClosed(()=>0,n=>throw new Exception("must not prompt"),()=>requests++,()=>Task.CompletedTask,1));Check(requests==0);
waits=0;running=1;
Check(await ExcelCloseGate.CloseNormally(()=>running,n=>true,()=>{},()=>{waits++;running=0;return Task.CompletedTask;},1)==ExcelCloseResult.Closed);Check(waits==1);
bool failed=false;
try{await ExcelCloseGate.ForceClosed(()=>throw new InvalidOperationException("inspection failed"),n=>true,()=>requests++,()=>Task.CompletedTask,1);}catch(InvalidOperationException){failed=true;}
Check(failed&&requests==0);
Console.WriteLine("PASS: 13 close gate cases — graceful close/cancel/timeout/final poll, force default refusal/success/still-running/no-process, inspection fails closed; no OS processes touched");
