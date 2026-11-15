using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Outline.Addin
{
    internal static class GkdAiStore
    {
        private const int MaxFileBytes = 1048576;
        private static readonly object Gate = new object();
        public static string PathName { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GKD", "ai-settings.json"); } }
        public static GkdAiSettings Load() { return Load(PathName, Unprotect); }
        public static void Save(GkdAiSettings settings) { Save(settings, PathName, Protect); }
        public static GkdAiProfile ResolveDefault() { return ResolveDefault(Load()); }
        public static GkdAiProfile ResolveDefault(GkdAiSettings settings)
        {
            ValidateSettings(settings);
            return settings.Profiles.Find(p => string.Equals(p.Id, settings.DefaultProfileId, StringComparison.OrdinalIgnoreCase)).Clone();
        }

        // The codec overloads are explicit test seams. Production never falls back to plaintext.
        public static GkdAiSettings Load(string path, Func<string, string> unprotect)
        {
            lock (Gate)
            {
                try
                {
                    if (unprotect == null) throw new InvalidDataException();
                    if (!File.Exists(path))
                    {
                        var placeholder = new GkdAiProfile { Name = "Dify（请配置密钥）" };
                        return new GkdAiSettings { Profiles = new List<GkdAiProfile> { placeholder }, DefaultProfileId = placeholder.Id };
                    }
                    string text;
                    using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        if (file.Length > MaxFileBytes) throw new InvalidDataException();
                        using (var reader = new StreamReader(file, new UTF8Encoding(false, true))) text = reader.ReadToEnd();
                    }
                    var envelope = GkdAiProtocol.ReadJson(text) as JObject;
                    if (envelope == null || envelope.Count != 3 || (int?)envelope["Version"] != 1 || envelope["Payload"] == null) throw new InvalidDataException();
                    var payload = envelope["Payload"] as JObject;
                    if (payload == null || !string.Equals((string)envelope["Checksum"], Digest(payload.ToString(Formatting.None)), StringComparison.Ordinal)) throw new InvalidDataException();
                    var settings = payload.ToObject<GkdAiSettings>();
                    if (settings == null || settings.Profiles == null || settings.Profiles.Count == 0 || settings.Profiles.Count > 100) throw new InvalidDataException();
                    foreach (var p in settings.Profiles)
                    {
                        if (p == null) throw new InvalidDataException();
                        if (!string.IsNullOrEmpty(p.ApiKey)) p.ApiKey = unprotect(p.ApiKey);
                    }
                    ValidateSettings(settings);
                    return settings;
                }
                catch (Exception e) when (Recoverable(e)) { throw new InvalidDataException("AI 配置读取失败：文件损坏、版本不支持或密钥无法解密，请重新配置。"); }
            }
        }

        public static void Save(GkdAiSettings settings, string path, Func<string, string> protect)
        {
            lock (Gate)
            {
                ValidateSettings(settings);
                string temp = null;
                try
                {
                    if (protect == null) throw new InvalidDataException();
                    var encrypted = new GkdAiSettings { DefaultProfileId = settings.DefaultProfileId };
                    foreach (var p in settings.Profiles)
                    {
                        var copy = p.Clone();
                        if (!string.IsNullOrEmpty(copy.ApiKey))
                        {
                            copy.ApiKey = protect(copy.ApiKey);
                            if (string.IsNullOrEmpty(copy.ApiKey)) throw new InvalidDataException();
                        }
                        encrypted.Profiles.Add(copy);
                    }
                    var payload = JObject.FromObject(encrypted);
                    string text = new JObject { ["Version"] = 1, ["Payload"] = payload, ["Checksum"] = Digest(payload.ToString(Formatting.None)) }.ToString(Formatting.None);
                    byte[] bytes = new UTF8Encoding(false).GetBytes(text);
                    if (bytes.Length > MaxFileBytes) throw new InvalidDataException();
                    string full = Path.GetFullPath(path);
                    Directory.CreateDirectory(Path.GetDirectoryName(full));
                    temp = full + "." + Guid.NewGuid().ToString("N") + ".tmp";
                    using (var file = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { file.Write(bytes, 0, bytes.Length); file.Flush(true); }
                    if (File.Exists(full)) File.Replace(temp, full, null);
                    else File.Move(temp, full);
                }
                catch (Exception e) when (Recoverable(e)) { throw new InvalidDataException("AI 配置保存失败，请检查目录权限和系统密钥保护功能。"); }
                finally { if (temp != null) { try { if (File.Exists(temp)) File.Delete(temp); } catch (IOException) { } catch (UnauthorizedAccessException) { } } }
            }
        }

        private static void ValidateSettings(GkdAiSettings settings)
        {
            if (settings == null || settings.Profiles == null || settings.Profiles.Count == 0 || settings.Profiles.Count > 100) throw new InvalidDataException("请保留1至100个 AI 配置。");
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in settings.Profiles)
            {
                if (p == null) throw new InvalidDataException("AI 配置内容无效。");
                p.ValidateConfiguration(false);
                if (!ids.Add(p.Id)) throw new InvalidDataException("AI 配置标识重复。");
            }
            if (settings.DefaultProfileId == null || !ids.Contains(settings.DefaultProfileId)) throw new InvalidDataException("请选择有效的默认 AI 配置。");
        }
        private static string Digest(string text)
        {
            using (var sha = SHA256.Create()) return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(text)));
        }
        private static bool Recoverable(Exception e) { return !(e is OutOfMemoryException) && !(e is StackOverflowException) && !(e is AccessViolationException); }
        private static string Protect(string value)
        {
#if NETFRAMEWORK
            return Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(value), null, DataProtectionScope.CurrentUser));
#else
            throw new PlatformNotSupportedException("生产配置密钥保护仅支持 Windows DPAPI。");
#endif
        }
        private static string Unprotect(string value)
        {
#if NETFRAMEWORK
            return Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(value), null, DataProtectionScope.CurrentUser));
#else
            throw new PlatformNotSupportedException("生产配置密钥保护仅支持 Windows DPAPI。");
#endif
        }
    }
}
