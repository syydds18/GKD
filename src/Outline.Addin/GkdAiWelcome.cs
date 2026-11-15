using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace Outline.Addin
{
    // A bare name is not a UDF call. SheetChange rewrites only the edited cell's
    // exact =GKD formula to =GKD(); the UDF supplies the welcome result.
    internal static class GkdAiWelcome
    {
        // Official PIA metadata (GUID and SheetChange DispId 1564 / 0x61c):
        // https://github.com/MicrosoftDocs/office-developer-excel-pia-ref-dotnet/blob/live/officedocs-dev-excel-pia/xml/Microsoft.Office.Interop.Excel/AppEvents.xml
        private static readonly Guid AppEvents = new Guid("00024413-0000-0000-C000-000000000046");
        private const int SheetChangeId = 0x61c;
        private static Subscription Active;
        private sealed class Subscription
        {
            internal object Application;
            internal Action<object, Guid, int, Delegate> Remove;
            internal Action<object, object> Handler;
            internal bool Live;
        }

        // AutoOpen must queue Start as an Excel macro; AutoClose calls Stop on the
        // same Excel main thread. Never attach from a UDF or a worker thread.
        internal static void Start()
        {
            if (Active != null) return;
            try
            {
                Start(ExcelDnaUtil.Application,
                    (app, iid, id, handler) => ComEventsHelper.Combine(app, iid, id, handler),
                    (app, iid, id, handler) => { ComEventsHelper.Remove(app, iid, id, handler); });
            }
            catch { /* Unavailable COM host must not prevent add-in startup. */ }
        }
        // Injectable COM transport lets portable tests exercise the real lifecycle.
        internal static void Start(object application, Action<object, Guid, int, Delegate> add, Action<object, Guid, int, Delegate> remove)
        {
            if (Active != null || application == null || add == null || remove == null) return;
            var subscription = new Subscription { Application = application, Remove = remove };
            subscription.Handler = (sheet, target) =>
            {
                if (subscription.Live && object.ReferenceEquals(Active, subscription))
                    HandleChange(subscription.Application, sheet, target);
            };
            try
            {
                add(application, AppEvents, SheetChangeId, subscription.Handler);
                subscription.Live = true;
                Active = subscription;
            }
            catch { Detach(subscription); }
        }
        internal static void Stop()
        {
            Subscription subscription = Active;
            Active = null;
            if (subscription != null) Detach(subscription);
        }
        private static void Detach(Subscription subscription)
        {
            subscription.Live = false;
            try { subscription.Remove(subscription.Application, AppEvents, SheetChangeId, subscription.Handler); }
            catch { /* Disconnected Excel or an incomplete connection point. */ }
            // Do not ReleaseComObject: Excel-DNA and other add-in code share the RCW.
            subscription.Application = null;
            subscription.Remove = null;
        }
        private static bool HasDefinedName(dynamic sheet)
        {
            // Two indexed lookups in this sheet/workbook only; never enumerate names
            // or create a name (which would shadow the registered Excel-DNA UDF).
            return NameExists(sheet.Names) || NameExists(sheet.Parent.Names);
        }
        private static bool NameExists(dynamic names)
        {
            try { return names.Item("GKD") != null; }
            catch (COMException e)
            {
                if (e.ErrorCode == unchecked((int)0x800A03EC)) return false; // Excel: name not found.
                throw; // Uncertain lookup => fail closed at the event boundary.
            }
        }
        private static bool Handling;
        internal static void HandleChange(object application, object sheet, object target)
        {
            if (Handling || application == null || sheet == null || target == null) return;
            dynamic app = application, range = target, worksheet = sheet;
            bool previous = false, restore = false;
            Handling = true;
            try
            {
                if ((bool)worksheet.ProtectContents || HasDefinedName(worksheet)) return;
                double count = Convert.ToDouble(range.CountLarge);
                if (double.IsNaN(count) || count < 1 || count > 10000 || count != Math.Truncate(count)) return;
                dynamic areas = range.Areas;
                int areaCount = (int)areas.Count;
                if (areaCount < 1 || areaCount > count) return;
                int visited = 0;
                for (int a = 1; a <= areaCount; a++)
                {
                    dynamic area = areas[a];
                    int cells = Convert.ToInt32(area.CountLarge);
                    if (cells < 1 || cells > count - visited) return;
                    for (int i = 1; i <= cells; i++)
                    {
                        visited++;
                        try
                        {
                            dynamic cell = area.Cells[i];
                            if (!(bool)cell.HasFormula || (bool)cell.MergeCells || (bool)cell.HasArray) continue;
                            if (!string.Equals((cell.Formula as string ?? "").Trim(), "=GKD", StringComparison.OrdinalIgnoreCase)) continue;
                            try { if ((bool)cell.HasSpill) continue; }
                            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { /* Pre-dynamic-array Excel. */ }
                            catch (System.Runtime.InteropServices.COMException e)
                            {
                                if (e.ErrorCode != unchecked((int)0x80020003) && e.ErrorCode != unchecked((int)0x80020006)) continue;
                            }
                            if (!restore)
                            {
                                previous = (bool)app.EnableEvents;
                                restore = true; // A failed setter may already have changed Excel's state.
                                try { app.EnableEvents = false; }
                                catch { return; }
                            }
                            // Only this exact edited cell is written; never workbook names or selection.
                            cell.Formula = "=GKD()";
                        }
                        catch { /* An inaccessible cell must not block other changed cells. */ }
                    }
                }
            }
            catch { /* Chart sheets, invalid targets, or Excel shutting down. */ }
            finally
            {
                if (restore) { try { app.EnableEvents = previous; } catch { /* Host no longer available. */ } }
                Handling = false;
            }
        }
    }
}
