using System;
using System.Linq;
using Outline.Setup;

internal static class DiscoveryTests
{
    internal static void Run()
    {
        Equal(@"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            ExcelPathCandidates.Executable("\"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE\" /dde"));
        Equal(@"D:\团队 工具\Office15\EXCEL.EXE",
            ExcelPathCandidates.Executable(@"D:\团队 工具\Office15\EXCEL.EXE /automation"));
        Equal(@"C:\Office\EXCEL.EXE", ExcelPathCandidates.Executable(@"C:\Office\EXCEL.EXE,0"));
        Equal(null, ExcelPathCandidates.Executable(@"C:\Office\EXCEL.EXE.bak"));
        Equal(null, ExcelPathCandidates.Executable(@"C:\Office\notexcel.exe /dde"));
        Equal(null, ExcelPathCandidates.Executable(@"cmd.exe /c C:\Office\EXCEL.EXE"));
        Equal(null, ExcelPathCandidates.Executable("\"C:\\Office\\EXCEL.EXE /dde"));
        Equal(null, ExcelPathCandidates.Executable(@"Office16\EXCEL.EXE"));
        Equal(null, ExcelPathCandidates.Executable(null));
        Environment.SetEnvironmentVariable("OUTLINE_TEST_OFFICE", @"D:\办公软件");
        try { Equal(@"D:\办公软件\EXCEL.EXE", ExcelPathCandidates.Executable(@"%OUTLINE_TEST_OFFICE%\EXCEL.EXE /dde")); }
        finally { Environment.SetEnvironmentVariable("OUTLINE_TEST_OFFICE", null); }
        var roots = ExcelPathCandidates.InstallRoot("\"D:\\Office\\\"").ToArray();
        Check(roots.Contains(@"D:\Office\EXCEL.EXE"), "direct InstallRoot");
        Check(roots.Contains(@"D:\Office\root\Office16\EXCEL.EXE"), "C2R root");
        Check(roots.Contains(@"D:\Office\Office15\EXCEL.EXE"), "Office 2013");
        Check(roots.Contains(@"D:\Office\Office14\EXCEL.EXE"), "Office 2010");
        Check(roots.Length == roots.Distinct(StringComparer.OrdinalIgnoreCase).Count(), "deduplication");
        Check(!ExcelPathCandidates.InstallRoot("relative").Any(), "reject relative roots");
        Check(!ExcelPathCandidates.InstallRoot("\"").Any(), "reject malformed quote without aborting discovery");
        Equal(@"\\server\share\Office16\EXCEL.EXE", ExcelPathCandidates.Executable("\"\\\\server\\share\\Office16\\EXCEL.EXE\" /dde"));
        Check(ExcelPathCandidates.InstallRoot(@"C:\").Contains(@"C:\EXCEL.EXE"), "drive root");
        Console.WriteLine("PASS Excel candidate parsing: quoted/unquoted commands, icon suffix, environment expansion, Unicode, roots, malformed/non-Excel rejection");
    }
    private static void Equal(string expected, string actual)
    {
        if (expected != actual) throw new Exception("Discovery expected " + expected + "; actual " + actual);
    }
    private static void Check(bool value, string name) { if (!value) throw new Exception("Discovery: " + name); }
}
