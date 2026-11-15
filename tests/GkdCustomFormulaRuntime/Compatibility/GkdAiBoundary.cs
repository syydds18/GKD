using ExcelDna.Integration;
namespace Outline.Addin
{
    // AI scheduling and reference I/O belong to GkdAiFunctions. This compatibility
    // build uses the production signatures and actual Excel-DNA v1.9 types.
    public static class GkdAiFunctions
    {
        public static object Gkd(object promptValue, object data = null, object outputMode = null) { return null; }
        internal static object[,] ReadReference(ExcelReference reference) { return null; }
    }
}
