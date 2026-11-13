using System;

namespace Outline.Core
{
    public static class HydroRangeAddress
    {
        public static bool IsValid(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) return false;
            string value = address.Replace("$", "").Trim().TrimStart('=');
            int bang = value.LastIndexOf('!');
            if (bang >= 0) value = value.Substring(bang + 1);
            string[] parts = value.Split(':');
            if (parts.Length < 1 || parts.Length > 2) return false;
            int r1, c1, r2, c2;
            if (!TryCell(parts[0], out r1, out c1)) return false;
            if (parts.Length == 1) return true;
            return TryCell(parts[1], out r2, out c2) && r2 >= r1 && c2 >= c1;
        }

        private static bool TryCell(string text, out int row, out int column)
        {
            row = column = 0;
            if (string.IsNullOrEmpty(text)) return false;
            int i = 0;
            while (i < text.Length && char.IsLetter(text[i]))
            {
                char c = char.ToUpperInvariant(text[i]);
                if (c < 'A' || c > 'Z') return false;
                column = checked(column * 26 + c - 'A' + 1);
                i++;
            }
            if (i == 0 || column > 16384 || i == text.Length) return false;
            for (; i < text.Length; i++)
            {
                if (!char.IsDigit(text[i])) return false;
                row = checked(row * 10 + text[i] - '0');
            }
            return row >= 1 && row <= 1048576;
        }
    }
}
