using System;
using System.Collections.Generic;
namespace Outline.Setup
{
    public static class CoreUtil
    {
        public static string OpenValue(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.IndexOf('"') >= 0 || path.IndexOf('\r') >= 0 || path.IndexOf('\n') >= 0)
                throw new ArgumentException("加载项路径无效。");
            return "/R \"" + path + "\"";
        }
        public static string ReadPeArchitecture(System.IO.Stream stream)
        {
            using (var reader = new System.IO.BinaryReader(stream, System.Text.Encoding.UTF8, true))
            {
                if (stream.Length < 64 || reader.ReadUInt16() != 0x5a4d) throw new System.IO.InvalidDataException("不是有效的 Windows PE 文件。");
                stream.Position = 0x3c;
                int offset = reader.ReadInt32();
                if (offset < 64 || offset > stream.Length - 6) throw new System.IO.InvalidDataException("PE 头超出文件边界。");
                stream.Position = offset;
                if (reader.ReadUInt32() != 0x00004550) throw new System.IO.InvalidDataException("PE 签名无效。");
                ushort machine = reader.ReadUInt16();
                if (machine == 0x014c) return "x86";
                if (machine == 0x8664) return "x64";
                throw new System.IO.InvalidDataException("暂不支持此 Excel 架构（仅 x86/x64）。");
            }
        }
        public static string CheckedRelativePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.IndexOf(':') >= 0 || path.StartsWith("/") || path.StartsWith("\\"))
                throw new ArgumentException("必须使用安全的相对路径。");
            foreach (var part in path.Replace('\\', '/').Split('/'))
                if (part == ".." || part == "." || part.Length == 0) throw new ArgumentException("相对路径不能跳出安装目录。");
            return path;
        }
        public static bool OwnsValue(string expected, object actual)
        {
            return !string.IsNullOrEmpty(expected) && actual is string && string.Equals(expected, (string)actual, StringComparison.Ordinal);
        }
        public static string NextOpenSlot(IEnumerable<string> names)
        {
            var occupied = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < int.MaxValue; i++)
            {
                string candidate = i == 0 ? "OPEN" : "OPEN" + i.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (!occupied.Contains(candidate)) return candidate;
            }
            throw new InvalidOperationException("没有可用的 Excel OPEN 注册槽位。");
        }
    }
}
