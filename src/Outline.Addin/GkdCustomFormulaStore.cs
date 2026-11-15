using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Outline.Addin
{
    // Persistence deliberately has no Excel/UI side effects. Call Reload on the Excel
    // thread after a successful Save; callers must not report registration failures as save failures.
    internal static class GkdCustomFormulaStore
    {
        private const int MaxFileBytes = 1048576;
        private static readonly object Gate = new object();
        public static string PathName { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GKD", "custom-formulas.json"); } }
        public static List<GkdCustomFormula> Load() { return Load(PathName); }
        public static void Save(List<GkdCustomFormula> formulas) { Save(formulas, PathName); }
        internal static List<GkdCustomFormula> Load(string path)
        {
            lock (Gate)
            {
                // Do not hide access errors behind File.Exists's false result.
                try
                {
                    using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        if (stream.Length > MaxFileBytes) throw new InvalidDataException("公式文件不能超过 1 MB。");
                        using (var reader = new StreamReader(stream, new UTF8Encoding(false, true)))
                            return GkdCustomFormulaJson.Import(reader.ReadToEnd());
                    }
                }
                catch (FileNotFoundException) { return new List<GkdCustomFormula>(); }
                catch (DirectoryNotFoundException) { return new List<GkdCustomFormula>(); }
            }
        }
        internal static void Save(List<GkdCustomFormula> formulas, string path)
        {
            lock (Gate)
            {
                GkdCustomFormulaFunctions.ValidateDefinitions(formulas);
                byte[] bytes = new UTF8Encoding(false, true).GetBytes(GkdCustomFormulaJson.Export(formulas));
                if (bytes.Length > MaxFileBytes) throw new InvalidDataException("公式文件不能超过 1 MB。");
                string full = Path.GetFullPath(path), temporary = null;
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(full));
                    temporary = full + "." + Guid.NewGuid().ToString("N") + ".tmp";
                    using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
                    // Same-directory replacement is atomic; never delete the old file as a fallback.
                    if (File.Exists(full)) File.Replace(temporary, full, null);
                    else File.Move(temporary, full);
                }
                finally
                {
                    if (temporary != null)
                    {
                        try { File.Delete(temporary); }
                        catch (IOException) { }
                        catch (UnauthorizedAccessException) { }
                    }
                }
            }
        }
    }
}
