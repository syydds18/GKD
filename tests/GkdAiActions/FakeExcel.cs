using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public sealed class FakeApp
{
    public FakeRange Selection;
    public object StatusBar;
    public bool EnableEvents=true;
    public object CutCopyMode;
}
public sealed class FakeSheet { public bool ProtectContents; }
public sealed class FakeRange
{
    public int Row,Column;
    public object Formula,Value2;
    public object Formula2 { get { return Formula; } }
    public bool HasFormula { get { return Formula is string && ((string)Formula).StartsWith("="); } }
    public FakeRange ParentSpill,Spilling;
    public bool NullSpillParent, NullSpillingRange;
    public FakeRange SpillParent { get { if(NullSpillParent)return null;if(ParentSpill==null)throw new COMException("no spill");return ParentSpill; } }
    public FakeRange SpillingToRange { get { if(NullSpillingRange)return null;if(Spilling==null)throw new COMException("no spill");return Spilling; } }
    public long CountLarge=1;
    public object MergeCells=false,HasArray=false;
    public bool ThrowOnPaste;
    public FakeSheet Worksheet=new FakeSheet();
    public List<FakeRange> Areas { get;set; }
    public List<FakeRange> Cells { get;set; }
    public int Dirties,Calculates,Pastes,Copies;
    public string Address { get { return "R"+Row+"C"+Column+":"+CountLarge; } }
    public FakeRange(int row,int col,object formula,object value)
    {Row=row;Column=col;Formula=formula;Value2=value;Cells=new List<FakeRange>{this};}
    public static FakeRange Selection(params FakeRange[] cells)
    {var r=new FakeRange(1,1,null,null);r.CountLarge=cells.Length;r.Cells=new List<FakeRange>(cells);r.Areas=new List<FakeRange>{r};return r;}
    public void Dirty(){Dirties++;}
    public void Calculate(){Calculates++;}
    public void Copy(){Copies++;}
    public FakeApp App;
    public void PasteSpecial(int type,object operation,object blanks,object transpose){Program.Check(type==-4163,"native paste values");if(App!=null)Program.Check(!App.EnableEvents,"events off during write");if(ThrowOnPaste)throw new COMException("paste failed");Pastes++;Formula=null;}
}
