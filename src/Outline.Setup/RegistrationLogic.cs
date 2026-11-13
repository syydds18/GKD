using System;
using System.Collections.Generic;
namespace Outline.Setup
{
    public sealed class Registration
    {
        public string OfficeVersion { get; set; }
        public string View { get; set; }
        public string Name { get; set; }
        public string Expected { get; set; }
    }
    public interface IOptionStore
    {
        string[] Names();
        object Get(string name);
        void Set(string name, string value);
        void Delete(string name);
    }
    public static class RegistrationLogic
    {
        public static Registration Add(IOptionStore store, string version, string view, string expected, Action<Registration> journal)
        {
            var entry = new Registration { OfficeVersion = version, View = view, Name = CoreUtil.NextOpenSlot(store.Names()), Expected = expected };
            journal(entry); // Persist exact ownership BEFORE any registry mutation.
            if (store.Get(entry.Name) != null) throw new InvalidOperationException("注册槽位已被其他程序占用，请重试。");
            store.Set(entry.Name, expected);
            if (!CoreUtil.OwnsValue(expected, store.Get(entry.Name))) throw new InvalidOperationException("Excel 加载项注册后核对失败。");
            return entry;
        }
        public static bool Remove(IOptionStore store, Registration entry)
        {
            if (!CoreUtil.OwnsValue(entry.Expected, store.Get(entry.Name))) return false;
            store.Delete(entry.Name);
            if (store.Get(entry.Name) != null) throw new InvalidOperationException("Excel 加载项注册项删除后核对失败。");
            return true;
        }
    }
}
