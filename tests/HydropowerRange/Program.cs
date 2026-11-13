using System;
using Outline.Core;

static class Program
{
    static int Main()
    {
        try
        {
            Check(HydroRangeAddress.IsValid("B2:K120"), "rectangular range");
            Check(HydroRangeAddress.IsValid("$B$2:$K$120"), "absolute range");
            Check(HydroRangeAddress.IsValid("XFD1048576"), "last Excel cell");
            Check(!HydroRangeAddress.IsValid("B2:K"), "reject whole row-column mix");
            Check(!HydroRangeAddress.IsValid("B2:K120,A1"), "reject multiple areas");
            Check(!HydroRangeAddress.IsValid("K120:B2"), "reject reversed bounds");
            Check(!HydroRangeAddress.IsValid("XFE1"), "reject column beyond Excel bounds");
            Check(!HydroRangeAddress.IsValid("A1048577"), "reject row beyond Excel bounds");
            Check(HydroRangeAddress.IsValid("='Sheet 1'!$B$2:$K$120"), "sheet-qualified print area");
            Check(Math.Abs(RowHeightPlanner.WithVerticalPadding(15, 3.5) - 22) < 0.0001, "adds padding above minimum");
            Check(Math.Abs(RowHeightPlanner.WithVerticalPadding(10, 3.5) - 20) < 0.0001, "applies minimum height");
            Console.WriteLine("PASS hydropower range address");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine("FAIL hydropower range address: " + ex.Message); return 1; }
    }

    private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
}
