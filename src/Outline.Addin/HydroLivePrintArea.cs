using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Outline.Addin
{
    /// <summary>Reads current in-memory print settings; never saves or edits the source sheet.</summary>
    internal static class HydroLivePrintArea
    {
        // Call on Excel's UI/COM thread. Enabling PrintCommunication commits ALL cached
        // PageSetup commands in this application, not just commands for this worksheet.
        internal static string Read(dynamic app, dynamic sheet)
        {
            bool restoreDisabled = false;
            bool canReadPrintArea = true;
            try
            {
                try
                {
                    restoreDisabled = !Convert.ToBoolean(app.PrintCommunication, CultureInfo.InvariantCulture);
                    if (restoreDisabled)
                    {
                        // Set the restoration guard before the setter: it may throw after
                        // partially changing state (for example a printer/driver failure).
                        canReadPrintArea = false;
                        app.PrintCommunication = true;
                        canReadPrintArea = true;
                    }
                }
                catch { /* Older/unsupported hosts: read best-effort without toggling. */ }

                string sheetName = "";
                try { sheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture); } catch { }
                bool pageSetupRead = false;
                string printArea = "";
                if (canReadPrintArea)
                {
                    try
                    {
                        printArea = Convert.ToString(sheet.PageSetup.PrintArea, CultureInfo.InvariantCulture);
                        pageSetupRead = true;
                    }
                    catch { }
                }

                string address = Normalize(printArea, sheetName);
                if (address.Length != 0) return address;

                // A successful empty PageSetup.PrintArea means it was CLEARED. Do not
                // resurrect an old hidden name. A valid PageSetup always wins too.
                if (!pageSetupRead)
                {
                    try
                    {
                        dynamic name = sheet.Names.Item("Print_Area");
                        dynamic range = name.RefersToRange;
                        if (object.ReferenceEquals((object)range.Parent, (object)sheet))
                        {
                            address = Normalize(Convert.ToString(range.Address, CultureInfo.InvariantCulture), sheetName);
                            if (address.Length != 0) return address;
                        }
                    }
                    catch { /* Missing/non-range Print_Area is normal. */ }
                }
                try { return Normalize(Convert.ToString(sheet.UsedRange.Address, CultureInfo.InvariantCulture), sheetName); }
                catch { return ""; }
            }
            finally
            {
                // Never mask the read result with a restoration error. Actual host-level
                // restoration cannot be guaranteed if Excel rejects this setter.
                if (restoreDisabled) { try { app.PrintCommunication = false; } catch { } }
            }
        }

        private static readonly Regex A1Area = new Regex(
            @"^(?:[A-Za-z]{1,3}[1-9][0-9]{0,6}(?::[A-Za-z]{1,3}[1-9][0-9]{0,6})?|[A-Za-z]{1,3}:[A-Za-z]{1,3}|[1-9][0-9]{0,6}:[1-9][0-9]{0,6})$",
            RegexOptions.CultureInvariant);

        private static string Normalize(string address, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(address)) return "";
            string text = address.Trim();
            if (text.StartsWith("=", StringComparison.Ordinal)) text = text.Substring(1).Trim();
            var result = new List<string>();
            bool quoted = false;
            int start = 0;
            for (int i = 0; i <= text.Length; i++)
            {
                if (i < text.Length && text[i] == '\'')
                {
                    if (quoted && i + 1 < text.Length && text[i + 1] == '\'') { i++; continue; }
                    quoted = !quoted;
                }
                if (i != text.Length && (text[i] != ',' || quoted)) continue;
                if (quoted) return "";
                string part = text.Substring(start, i - start).Trim();
                // The last ! is the reference separator, even if the quoted worksheet
                // name itself contains !. Decode escaped apostrophes before matching.
                int bang = part.LastIndexOf('!');
                if (bang >= 0)
                {
                    string qualifier = part.Substring(0, bang).Trim();
                    if (qualifier.StartsWith("'", StringComparison.Ordinal) && qualifier.EndsWith("'", StringComparison.Ordinal) && qualifier.Length >= 2)
                        qualifier = qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'");
                    // Do not strip an arbitrary sheet/workbook prefix into a local range.
                    if (string.IsNullOrEmpty(sheetName) || !string.Equals(qualifier, sheetName, StringComparison.OrdinalIgnoreCase)) return "";
                    part = part.Substring(bang + 1).Trim();
                }
                part = part.Replace("$", "");
                if (!A1Area.IsMatch(part)) return "";
                result.Add(part);
                start = i + 1;
            }
            return string.Join(",", result);
        }
    }
}
