using System;
using Outline.Setup;
internal static class Program
{
    private static int Main()
    {
        try
        {
            DiscoveryTests.Run();
            Equal(true, ExcelProcessScope.Matches("S-1-5-21-1", 2, "S-1-5-21-1", 2));
            Equal(false, ExcelProcessScope.Matches("S-1-5-21-1", 2, "S-1-5-21-2", 2));
            Equal(false, ExcelProcessScope.Matches("S-1-5-21-1", 2, "S-1-5-21-1", 3));
            Equal(false, ExcelProcessScope.Matches(null, 2, null, 2));
            Console.WriteLine("PASS Excel process scope requires verified same user SID and session");
            Equal("OPEN2", CoreUtil.NextOpenSlot(new[] { "OPEN", "OPEN1", "OPEN3", "Other" }));
            Equal("OPEN", CoreUtil.NextOpenSlot(new[] { "OPEN1", "OPEN01" }));
            Equal("OPEN1", CoreUtil.NextOpenSlot(new[] { "open" }));
            Console.WriteLine("PASS OPEN slots preserve occupied values and use first free slot");
            Equal(true, CoreUtil.OwnsValue("/R \"C:\\工具\\Outline.xll\"", "/R \"C:\\工具\\Outline.xll\""));
            Equal(false, CoreUtil.OwnsValue("/R \"C:\\工具\\Outline.xll\"", "/R \"C:\\Other.xll\""));
            Equal(false, CoreUtil.OwnsValue("", ""));
            Equal(false, CoreUtil.OwnsValue("a", null));
            Equal(false, CoreUtil.OwnsValue("a", "A"));
            Console.WriteLine("PASS ownership requires exact persisted non-empty value");
            Equal("/R \"C:\\团队 工具\\x64\\Outline.xll\"", CoreUtil.OpenValue("C:\\团队 工具\\x64\\Outline.xll"));
            Equal("x86", CoreUtil.ReadPeArchitecture(Pe(0x014c)));
            Equal("x64", CoreUtil.ReadPeArchitecture(Pe(0x8664)));
            Throws(() => CoreUtil.ReadPeArchitecture(new System.IO.MemoryStream(new byte[8])));
            Throws(() => CoreUtil.ReadPeArchitecture(Pe(0xaa64)));
            Equal("x64/Outline.xll", CoreUtil.CheckedRelativePath("x64/Outline.xll"));
            Throws(() => CoreUtil.CheckedRelativePath("../outside"));
            Throws(() => CoreUtil.CheckedRelativePath("x64/../../outside"));
            Throws(() => CoreUtil.CheckedRelativePath("C:\\outside"));
            Throws(() => CoreUtil.CheckedRelativePath("/outside"));
            Console.WriteLine("PASS quoted OPEN command, PE bitness, path traversal rejection");
            var registry = new FakeOptions();
            registry.Set("OPEN", "other addin");
            Registration saved = null;
            var record = RegistrationLogic.Add(registry, "16.0", "Registry32", "our addin", x => {
                Equal(null, registry.Get(x.Name)); saved = x;
            });
            Equal("OPEN1", record.Name); Equal(record, saved); Equal("other addin", registry.Get("OPEN"));
            Equal(true, RegistrationLogic.Remove(registry, record)); Equal(null, registry.Get("OPEN1"));
            registry.Set("OPEN1", "user replacement");
            Equal(false, RegistrationLogic.Remove(registry, record)); Equal("user replacement", registry.Get("OPEN1"));
            Throws(() => RegistrationLogic.Add(registry, "16.0", "Registry32", "new addin", x => { throw new Exception("disk full"); }));
            Equal(null, registry.Get("OPEN2"));
            Console.WriteLine("PASS write-ahead ownership journal, exact removal, preserve replacements and other addins");
            var folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "OutlineSetupTests-" + Guid.NewGuid().ToString("N"));
            System.IO.Directory.CreateDirectory(folder);
            try
            {
                var path = System.IO.Path.Combine(folder, "state.xml");
                var state = new InstallState { Architecture = "x64", OfficeVersion = "16.0", UninstallCommand = "\"C:\\团队 工具\\Outline.Setup.exe\" /uninstall" };
                state.Files.Add("x64/Outline.xll"); state.Registrations.Add(record);
                StateFile.Save(path, state);
                var loaded = StateFile.Load(path);
                Equal("x64", loaded.Architecture); Equal("16.0", loaded.OfficeVersion);
                Equal(state.UninstallCommand, loaded.UninstallCommand); Equal("x64/Outline.xll", loaded.Files[0]);
                Equal(record.Name, loaded.Registrations[0].Name); Equal(record.Expected, loaded.Registrations[0].Expected); Equal(false, loaded.Complete);
                loaded.Complete = true; StateFile.Save(path, loaded); Equal(true, StateFile.Load(path).Complete);
                Equal(false, System.IO.File.Exists(path + ".new"));
                System.IO.File.WriteAllText(path, "<!DOCTYPE x [<!ENTITY evil SYSTEM 'file:///etc/passwd'>]><InstallState>&evil;</InstallState>");
                Throws(() => StateFile.Load(path));
            }
            finally { System.IO.Directory.Delete(folder, true); }
            Console.WriteLine("PASS XML state roundtrip, atomic replacement, non-ASCII paths, DTD rejection");
            return 0;
        }
        catch (Exception e) { Console.Error.WriteLine("FAIL " + e.Message); return 1; }
    }
    private sealed class FakeOptions : IOptionStore
    {
        private readonly System.Collections.Generic.Dictionary<string, string> data = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string[] Names() { return new System.Collections.Generic.List<string>(data.Keys).ToArray(); }
        public object Get(string name) { return data.TryGetValue(name, out var value) ? value : null; }
        public void Set(string name, string value) { data[name] = value; }
        public void Delete(string name) { data.Remove(name); }
    }
    private static System.IO.MemoryStream Pe(ushort machine)
    {
        var stream = new System.IO.MemoryStream(new byte[128]);
        var writer = new System.IO.BinaryWriter(stream);
        writer.Write((ushort)0x5a4d); stream.Position = 0x3c; writer.Write(0x40);
        stream.Position = 0x40; writer.Write(0x00004550); writer.Write(machine); stream.Position = 0;
        return stream;
    }
    private static void Throws(Action action)
    {
        try { action(); } catch (Exception) { return; }
        throw new Exception("Expected invalid input to be rejected");
    }
    private static void Equal<T>(T expected, T actual)
    {
        if (!object.Equals(expected, actual)) throw new Exception("Expected: " + expected + "; actual: " + actual);
    }
}
