using System;

namespace Outline.Core
{
    public static class RowHeightPlanner
    {
        public static double WithVerticalPadding(double fittedHeight, double paddingEachSide, double minimum = 20.0)
        {
            if (double.IsNaN(fittedHeight) || double.IsInfinity(fittedHeight) || fittedHeight < 0) throw new ArgumentOutOfRangeException("fittedHeight");
            if (double.IsNaN(paddingEachSide) || double.IsInfinity(paddingEachSide) || paddingEachSide < 0) throw new ArgumentOutOfRangeException("paddingEachSide");
            if (double.IsNaN(minimum) || double.IsInfinity(minimum) || minimum < 0) throw new ArgumentOutOfRangeException("minimum");
            return Math.Max(minimum, fittedHeight + 2 * paddingEachSide);
        }
    }
}
