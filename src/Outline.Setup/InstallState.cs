using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace Outline.Setup
{
    public sealed class InstallState
    {
        public string ProductId { get; set; } = "ExcelOutlineAddin/1.0";
        public string Architecture { get; set; }
        public string OfficeVersion { get; set; }
        public bool Complete { get; set; }
        public List<Registration> Registrations { get; set; } = new List<Registration>();
        public List<string> Files { get; set; } = new List<string>();
        public string UninstallCommand { get; set; }
    }
    internal static class StateFile
    {
        internal static InstallState Load(string path)
        {
            using (var reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1024 * 1024 }))
                return (InstallState)new XmlSerializer(typeof(InstallState)).Deserialize(reader);
        }
        internal static void Save(string path, InstallState state)
        {
            string temp = path + ".new";
            using (var stream = new FileStream(temp, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                new XmlSerializer(typeof(InstallState)).Serialize(stream, state);
                stream.Flush(true);
            }
            if (File.Exists(path)) File.Replace(temp, path, null);
            else File.Move(temp, path);
            var check = Load(path);
            if (!Same(state, check)) throw new IOException("安装状态保存后核对失败。");
        }
        private static bool Same(InstallState expected, InstallState actual)
        {
            if (actual == null || expected.ProductId != actual.ProductId || expected.Architecture != actual.Architecture || expected.OfficeVersion != actual.OfficeVersion || expected.Complete != actual.Complete || expected.UninstallCommand != actual.UninstallCommand || expected.Registrations.Count != actual.Registrations.Count || expected.Files.Count != actual.Files.Count) return false;
            for (int i = 0; i < expected.Files.Count; i++) if (expected.Files[i] != actual.Files[i]) return false;
            for (int i = 0; i < expected.Registrations.Count; i++)
            {
                var e = expected.Registrations[i]; var a = actual.Registrations[i];
                if (e.OfficeVersion != a.OfficeVersion || e.View != a.View || e.Name != a.Name || e.Expected != a.Expected) return false;
            }
            return true;
        }
    }
}
