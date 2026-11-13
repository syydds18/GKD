using System;
using System.Collections.Generic;
using System.Dynamic;
using Outline.Addin;

static class Program
{
    static int Main()
    {
        RefreshesAfterEnablingPrintCommunication();
        PreservesEnabledPrintCommunication();
        FallsBackAndRestoresOnReadFailure();
        RestoresWhenEnableFails();
        UsesFreshNamesOnlyWhenPageSetupCannotBeRead();
        EmptyPrintAreaDoesNotResurrectStaleName();
        HandlesQualifiedAddresses();
        DoesNotSaveOrChangeSavedFlag();
        RestoresAfterPartialEnableFailure();
        ToleratesUnsupportedPrintCommunication();
        RestoresAfterAllRangeReadsFail();
        Console.WriteLine("PASS HydroLivePrintArea dynamic COM contract tests");
        return 0;
    }

    static void RefreshesAfterEnablingPrintCommunication()
    {
        var app = new FakeApp(false);
        var page = new FakePage("$A$1:$C$9", app.Events);
        app.OnEnable = page.RecordCommit;
        string first = HydroLivePrintArea.Read(app, new FakeSheet(page, "$A$1:$Z$50"));
        page.PendingValue = "$B$2:$D$11";
        string second = HydroLivePrintArea.Read(app, new FakeSheet(page, "$A$1:$Z$50"));
        Equal("A1:C9", first, "first area");
        Equal("B2:D11", second, "fresh area on each read");
        Equal(false, app.PrintCommunication, "original disabled state restored");
        Equal("set:true,get:print-area,set:false,set:true,get:print-area,set:false", string.Join(",", app.Events), "flush/read/restore order");
    }

    static void PreservesEnabledPrintCommunication()
    {
        var app = new FakeApp(true);
        string address = HydroLivePrintArea.Read(app, new FakeSheet(new FakePage("$A$1:$A$2", app.Events), "$A$1:$X$20"));
        Equal("A1:A2", address, "area");
        Equal("get:print-area", string.Join(",", app.Events), "no unnecessary toggle");
        Equal(true, app.PrintCommunication, "enabled state remains enabled");
    }

    static void FallsBackAndRestoresOnReadFailure()
    {
        var app = new FakeApp(false);
        var page = new FakePage(null, app.Events) { ThrowOnRead = true };
        string address = HydroLivePrintArea.Read(app, new FakeSheet(page, "$C$3:$H$17"));
        Equal("C3:H17", address, "used range fallback");
        Equal(false, app.PrintCommunication, "state restored after PageSetup failure");
    }

    static void RestoresWhenEnableFails()
    {
        var app = new FakeApp(false) { ThrowOnEnable = true };
        string address = HydroLivePrintArea.Read(app, new FakeSheet(new FakePage("$A$1:$B$3", app.Events), "$D$4:$E$8"));
        Equal("D4:E8", address, "fallback when flush cannot be enabled");
        Equal(false, app.PrintCommunication, "disabled state restored after enable failure");
    }

    static void UsesFreshNamesOnlyWhenPageSetupCannotBeRead()
    {
        var app = new FakeApp(true);
        var page = new FakePage("$A$1:$B$2", app.Events);
        var sheet = new FakeSheet(page, "$A$1:$Z$90");
        sheet.NamedAddress = "$C$3:$D$4";
        Equal("A1:B2", HydroLivePrintArea.Read(app, sheet), "valid PageSetup wins over conflicting name");
        Equal(0, sheet.NameReads, "no named lookup needed for valid PageSetup");
        page.ThrowOnRead = true;
        Equal("C3:D4", HydroLivePrintArea.Read(app, sheet), "name fallback after failed PageSetup");
        sheet.NamedAddress = "$E$5:$F$6";
        Equal("E5:F6", HydroLivePrintArea.Read(app, sheet), "fresh name every read");
        sheet.ForeignNamedParent = new object();
        Equal("A1:Z90", HydroLivePrintArea.Read(app, sheet), "reject foreign named parent");
    }

    static void EmptyPrintAreaDoesNotResurrectStaleName()
    {
        var app = new FakeApp(true);
        var sheet = new FakeSheet(new FakePage("", app.Events), "$B$2:$G$20");
        sheet.NamedAddress = "$A$1:$B$3";
        Equal("B2:G20", HydroLivePrintArea.Read(app, sheet), "cleared area uses UsedRange");
        Equal(0, sheet.NameReads, "empty successful PageSetup is authoritative");
    }

    static void HandlesQualifiedAddresses()
    {
        var app = new FakeApp(true);
        var page = new FakePage("='水电, O''Brien!$'!$B$2:$D$9,'水电, O''Brien!$'!$F$2:$H$4", app.Events);
        var sheet = new FakeSheet(page, "$A$1:$Z$30") { SheetName = "水电, O'Brien!$" };
        Equal("B2:D9,F2:H4", HydroLivePrintArea.Read(app, sheet), "quoted comma/apostrophe/exclamation sheet qualifier");
        page.Value = "='Other sheet'!$A$1:$B$3";
        Equal("A1:Z30", HydroLivePrintArea.Read(app, sheet), "reject wrong worksheet qualifier");
        page.Value = "='[Other.xlsx]水电, O''Brien!$'!$A$1:$B$3";
        Equal("A1:Z30", HydroLivePrintArea.Read(app, sheet), "reject external qualifier");
    }

    static void DoesNotSaveOrChangeSavedFlag()
    {
        foreach (bool saved in new[] { false, true })
        {
            var app = new FakeApp(false);
            var sheet = new FakeSheet(new FakePage("$A$1:$C$8", app.Events), "$A$1:$Z$40");
            sheet.Book.Saved = saved;
            HydroLivePrintArea.Read(app, sheet);
            Equal(saved, sheet.Book.Saved, "Saved state untouched");
            Equal(0, sheet.Book.SaveCalls, "no save call");
            Equal(0, sheet.ForbiddenOperations, "no source writes or reset breaks");
        }
    }

    static void RestoresAfterPartialEnableFailure()
    {
        var app = new FakeApp(false) { ThrowAfterEnable = true };
        var sheet = new FakeSheet(new FakePage("$A$1:$C$8", app.Events), "$A$1:$Z$40");
        Equal("A1:Z40", HydroLivePrintArea.Read(app, sheet), "partial setter failure uses fallback");
        Equal(false, app.PrintCommunication, "restored after partially successful setter");
    }

    static void ToleratesUnsupportedPrintCommunication()
    {
        var app = new FakeApp(true) { ThrowOnGet = true };
        var sheet = new FakeSheet(new FakePage("$C$3:$D$8", app.Events), "$A$1:$Z$40");
        Equal("C3:D8", HydroLivePrintArea.Read(app, sheet), "best-effort PageSetup on unsupported host");
        Equal("get:print-area", string.Join(",", app.Events), "no un-restorable state toggle");
    }

    static void RestoresAfterAllRangeReadsFail()
    {
        var app = new FakeApp(false);
        var sheet = new FakeSheet(new FakePage(null, app.Events) { ThrowOnRead = true }, null) { ThrowOnUsedRange = true };
        Equal("", HydroLivePrintArea.Read(app, sheet), "empty when all reads fail");
        Equal(false, app.PrintCommunication, "restored after all fallbacks fail");
    }

    static void Equal<T>(T expected, T actual, string message)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new Exception(message + ": expected " + expected + ", got " + actual);
    }

    sealed class FakeApp : DynamicObject
    {
        bool _printCommunication;
        Action _onEnable;
        public readonly List<string> Events = new List<string>();
        public bool ThrowOnEnable;
        public bool ThrowAfterEnable;
        public bool ThrowOnGet;
        public FakeApp(bool enabled) { _printCommunication = enabled; }
        public Action OnEnable { set { _onEnable = value; } }
        public bool PrintCommunication { get { return _printCommunication; } }
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            if (binder.Name == "PrintCommunication")
            {
                if (ThrowOnGet) throw new InvalidOperationException("unsupported");
                result = _printCommunication; return true;
            }
            result = null; return false;
        }
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            if (binder.Name != "PrintCommunication") return false;
            bool enabled = Convert.ToBoolean(value);
            Events.Add("set:" + (enabled ? "true" : "false"));
            if (enabled && ThrowOnEnable) throw new InvalidOperationException("enable failed");
            _printCommunication = enabled;
            if (enabled && ThrowAfterEnable) throw new InvalidOperationException("driver failure after setter");
            if (enabled && _onEnable != null) _onEnable();
            return true;
        }
    }

    sealed class FakePage : DynamicObject
    {
        readonly List<string> _events;
        public string Value;
        public string PendingValue;
        public bool ThrowOnRead;
        public FakePage(string value, List<string> events) { Value = value; _events = events; }
        public void RecordCommit() { if (PendingValue != null) { Value = PendingValue; PendingValue = null; } }
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            if (binder.Name != "PrintArea") { result = null; return false; }
            _events.Add("get:print-area");
            if (ThrowOnRead) throw new InvalidOperationException("PageSetup read failed");
            result = Value; return true;
        }
    }

    sealed class FakeSheet : DynamicObject
    {
        readonly FakePage _page;
        readonly string _usedRange;
        public string SheetName = "Sheet1";
        public string NamedAddress;
        public object ForeignNamedParent;
        public int NameReads;
        public int ForbiddenOperations;
        public bool ThrowOnUsedRange;
        public readonly FakeWorkbook Book = new FakeWorkbook();
        public FakeSheet(FakePage page, string usedRange) { _page = page; _usedRange = usedRange; }
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            if (binder.Name == "PageSetup") { result = _page; return true; }
            if (binder.Name == "UsedRange")
            {
                if (ThrowOnUsedRange) throw new InvalidOperationException("UsedRange unavailable");
                result = new FakeRange(_usedRange); return true;
            }
            if (binder.Name == "Name") { result = SheetName; return true; }
            if (binder.Name == "Parent") { result = Book; return true; }
            if (binder.Name == "Names") { NameReads++; result = new FakeNames(this); return true; }
            result = null; return false;
        }
        public override bool TrySetMember(SetMemberBinder binder, object value)
        { ForbiddenOperations++; throw new Exception("source must be read-only"); }
        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        { ForbiddenOperations++; throw new Exception("unexpected source method " + binder.Name); }
    }

    sealed class FakeNames : DynamicObject
    {
        readonly FakeSheet _sheet;
        public FakeNames(FakeSheet sheet) { _sheet = sheet; }
        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            if (binder.Name != "Item" || args.Length != 1 || (string)args[0] != "Print_Area")
                throw new Exception("unexpected Names lookup");
            if (_sheet.NamedAddress == null) throw new InvalidOperationException("missing name");
            dynamic name = new ExpandoObject();
            name.RefersToRange = new FakeRange(_sheet.NamedAddress, _sheet.ForeignNamedParent ?? _sheet);
            result = name; return true;
        }
    }

    sealed class FakeWorkbook
    {
        public bool Saved;
        public int SaveCalls;
        public void Save() { SaveCalls++; }
        public void SaveAs(object path) { SaveCalls++; }
    }

    sealed class FakeRange : DynamicObject
    {
        readonly string _address;
        readonly object _parent;
        public FakeRange(string address, object parent = null) { _address = address; _parent = parent; }
        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            if (binder.Name == "Address") { result = _address; return true; }
            if (binder.Name == "Parent") { result = _parent; return true; }
            result = null; return false;
        }
    }
}