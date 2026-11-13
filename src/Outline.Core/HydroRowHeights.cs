using System;

namespace Outline.Core
{
    public static class HydroRowHeights
    {
        // Excel returns null for mixed row heights. Probe homogeneous spans,
        // splitting only mixed spans; keep original per-row padding semantics.
        public static void Pad(int count, Func<int, int, double?> read, Action<int, int, double> write)
        {
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            if (read == null || write == null) throw new ArgumentNullException();
            if (count > 0) Apply(1, count, read, write);
        }

        private static void Apply(int first, int last, Func<int, int, double?> read, Action<int, int, double> write)
        {
            double? height = read(first, last);
            if (height.HasValue)
            {
                write(first, last, Math.Min(409.0, Math.Max(20.0, height.Value + 7.0)));
                return;
            }
            if (first == last) throw new InvalidOperationException("无法读取导出行高。");
            int middle = first + (last - first) / 2;
            Apply(first, middle, read, write);
            Apply(middle + 1, last, read, write);
        }
    }
}
