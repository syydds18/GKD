using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Outline.Addin;
class Program
{
 static void Main()
 {
  var book=new Book();var a=book.Add("A");var b=book.Add("B");var c=book.Add("C");book.Selected.AddRange(new[]{a,b});book.Active=b;
  using(var scope=new HydroSheetSelectionScope(book.Application,book))
  {
   scope.SelectOnly(a);
   Check(book.Selected.SequenceEqual(new[]{a}),"group selection must be removed before source copy");
   scope.SelectOnly(c);
   Check(book.Selected.SequenceEqual(new[]{c}),"each copy must isolate its own sheet");
  }
  Check(book.Selected.SequenceEqual(new[]{a,b})&&book.Active==b,"original selected tabs and active tab restored");
  try{using(var scope=new HydroSheetSelectionScope(book.Application,book)){scope.SelectOnly(c);throw new InvalidOperationException();}}catch(InvalidOperationException){}
  Check(book.Selected.SequenceEqual(new[]{a,b})&&book.Active==b,"restore grouped tabs on failure");
  var printer=new Printer();using(var batch=new HydroPrintCommunicationScope(printer)){Check(!printer.PrintCommunication,"page changes queued");}Check(printer.PrintCommunication,"printer restored");
  printer.PrintCommunication=false;using(var batch=new HydroPrintCommunicationScope(printer)){}Check(!printer.PrintCommunication&&printer.CommitCount==2,"commit queued export settings while restoring disabled caller state");
  var appearance=new AppearanceSheet();HydroExportAppearance.ClearFills(appearance);
  Check(appearance.Cells.FormatConditions.Deleted&&appearance.Cells.Interior.Pattern==-4142&&appearance.Cells.Interior.ColorIndex==-4142,"exported cells must have no direct or conditional fill");
  Console.WriteLine("PASS grouped-sheet isolation/restoration and no-fill appearance (portable model)");
 }
 static void Check(bool b,string s){if(!b)throw new Exception(s);}
}
public class Printer {bool enabled=true;public int CommitCount;public bool PrintCommunication {get{return enabled;}set{if(value)CommitCount++;enabled=value;}}}
public class AppearanceSheet {public AppearanceCells Cells {get;}=new AppearanceCells();}
public class AppearanceCells {public Conditions FormatConditions {get;}=new Conditions();public Interior Interior {get;}=new Interior();}
public class Conditions {public bool Deleted;public void Delete(){Deleted=true;}}
public class Interior {public int Pattern{get;set;}public int ColorIndex{get;set;}}
public class Book
{
 public List<Sheet> Selected=new List<Sheet>(); public Sheet Active;public object Application {get{return this;}} public object ActiveWindow {get{return this;}} public object SelectedSheets {get{return Selected;}} public object ActiveSheet {get{return Active;}}
 public Sheet Add(string n){return new Sheet{Parent=this,Name=n};}
 public void Activate(){}
}
public class Sheet
{
 public Book Parent; public string Name;
 public void Select(object Replace){if(Convert.ToBoolean(Replace))Parent.Selected.Clear();if(!Parent.Selected.Contains(this))Parent.Selected.Add(this);Parent.Active=this;}
 public void Activate(){Parent.Active=this;}
}
