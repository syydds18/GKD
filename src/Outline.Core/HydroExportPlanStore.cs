using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace Outline.Core
{
    public sealed class HydroExportPlan
    {
        public HydroExportPlan()
        {
            Sheets = new List<HydroExportPlanSheet>();
        }

        public List<HydroExportPlanSheet> Sheets { get; private set; }
    }

    public sealed class HydroExportPlanSheet
    {
        public string Name { get; set; }
        public string PrintRange { get; set; }
    }

    public static class HydroExportPlanPolicy
    {
        public static bool ShouldInclude(string sheetName, bool recommended,
            IEnumerable<string> currentSelection, HydroExportPlan savedPlan)
        {
            var current = currentSelection == null ? null : currentSelection.ToList();
            if (current != null && current.Count > 0)
                return current.Any(x => string.Equals(x, sheetName, StringComparison.OrdinalIgnoreCase));
            return recommended;
        }

        // Only an explicit load uses saved ranges. No matches means the caller must
        // leave its current choices untouched (for example after sheets were renamed).
        public static List<HydroExportPlanSheet> MatchVisibleRanges(HydroExportPlan savedPlan,
            IEnumerable<string> visibleSheetNames)
        {
            var matches = new List<HydroExportPlanSheet>();
            if (savedPlan == null || visibleSheetNames == null) return matches;
            foreach (string name in visibleSheetNames.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var saved = savedPlan.Sheets.FirstOrDefault(x => x != null &&
                    string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase) && ValidRangeForSheet(x.PrintRange, name));
                if (saved != null) matches.Add(new HydroExportPlanSheet { Name = name, PrintRange = saved.PrintRange });
            }
            return matches;
        }

        private static bool ValidRangeForSheet(string address, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(address)) return false;
            string value = address.Trim().TrimStart('=');
            int bang = value.LastIndexOf('!');
            if (bang >= 0)
            {
                string qualifier = value.Substring(0, bang).Trim();
                if (qualifier.Length >= 2 && qualifier.StartsWith("'", StringComparison.Ordinal) && qualifier.EndsWith("'", StringComparison.Ordinal))
                    qualifier = qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'");
                if (!string.Equals(qualifier, sheetName, StringComparison.OrdinalIgnoreCase)) return false;
            }
            try { return HydroRangeAddress.IsValid(address); }
            catch (OverflowException) { return false; }
        }
    }

    public sealed class HydroExportPlanStore
    {
        private readonly string directory;

        public HydroExportPlanStore(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("A plan directory is required.", "directory");
            this.directory = directory;
        }

        public static HydroExportPlanStore ForLocalApplicationData()
        {
            string root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return new HydroExportPlanStore(Path.Combine(root, "Outline", "HydroExportPlans"));
        }

        public HydroExportPlan Load(string workbookPath)
        {
            string path = GetPlanPath(workbookPath);
            if (!File.Exists(path)) return null;
            try
            {
                var document = XDocument.Load(path);
                if (document.Root == null || document.Root.Name != "HydroExportPlan" ||
                    (string)document.Root.Attribute("version") != "1") return null;
                var plan = new HydroExportPlan();
                foreach (var element in document.Root.Elements("Sheet"))
                {
                    string name = (string)element.Attribute("name");
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    plan.Sheets.Add(new HydroExportPlanSheet { Name = name, PrintRange = (string)element.Attribute("range") ?? "" });
                }
                return plan;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is System.Xml.XmlException)
            {
                return null;
            }
        }

        public void Save(string workbookPath, IEnumerable<HydroExportPlanSheet> sheets)
        {
            if (sheets == null) throw new ArgumentNullException("sheets");
            var items = sheets.Where(x => x != null && !string.IsNullOrWhiteSpace(x.Name)).ToList();
            if (items.Count == 0) throw new ArgumentException("At least one sheet must be saved.", "sheets");
            string path = GetPlanPath(workbookPath);
            Directory.CreateDirectory(directory);
            string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                new XDocument(new XElement("HydroExportPlan", new XAttribute("version", "1"),
                    items.Select(x => new XElement("Sheet", new XAttribute("name", x.Name), new XAttribute("range", x.PrintRange ?? ""))))).Save(temp);
                if (File.Exists(path))
                {
                    try { File.Replace(temp, path, null); }
                    catch (PlatformNotSupportedException) { File.Delete(path); File.Move(temp, path); }
                    catch (IOException) { File.Delete(path); File.Move(temp, path); }
                }
                else File.Move(temp, path);
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }

        private string GetPlanPath(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath)) throw new ArgumentException("A workbook path is required.", "workbookPath");
            string canonical = Path.GetFullPath(workbookPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            byte[] bytes;
            using (var sha = SHA256.Create()) bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToUpperInvariant()));
            return Path.Combine(directory, BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant() + ".xml");
        }
    }
}