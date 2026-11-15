using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Newtonsoft.Json.Linq;

namespace Outline.Addin
{
    // Diagnostics use structural metadata only. Never write request/response bodies,
    // exception messages, credentials, workbook identities or selected cell values.
    internal static class GkdAiDiagnostics
    {
        private static readonly object Gate = new object();
        internal static string LogDirectory { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelOutlineAddin", "logs"); } }
        internal static string LogPath { get { return Path.Combine(LogDirectory, "ai-diagnostics.log"); } }
        internal static string Write(string id, string stage, string details)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(LogDirectory);
                    if (File.Exists(LogPath) && new FileInfo(LogPath).Length > 2 * 1024 * 1024)
                    {
                        string previous = LogPath + ".1";
                        if (File.Exists(previous)) File.Delete(previous);
                        File.Move(LogPath, previous);
                    }
                    File.AppendAllText(LogPath, DateTimeOffset.Now.ToString("o") + " request=" + id + " stage=" + stage + " " + details + Environment.NewLine, new UTF8Encoding(false));
                }
                return LogPath;
            }
            catch (Exception) { return null; } // Logging failures must not change requests.
        }
        internal static string EnvironmentSummary(HttpWebRequest request)
        {
            var endpoint = request.RequestUri;
            // Host only; path, query, user name and proxy URL may contain private values.
            return "addin=" + typeof(GkdAiDiagnostics).Assembly.GetName().Version + " os=" + Environment.OSVersion.VersionString
                + " clr=" + Environment.Version + " process_bits=" + (IntPtr.Size * 8)
                + " tls_policy=" + ServicePointManager.SecurityProtocol
                + " strong_crypto_disabled=" + Switch("Switch.System.Net.DontEnableSchUseStrongCrypto")
                + " system_tls_disabled=" + Switch("Switch.System.Net.DontEnableSystemDefaultTlsVersions")
                + " host=" + endpoint.DnsSafeHost + " port=" + endpoint.Port
                + " proxy=" + (request.Proxy == null ? "none" : request.Proxy.GetType().Name)
                + " timeout_ms=" + request.Timeout;
        }
        private static string Switch(string name)
        {
            bool value;
            return AppContext.TryGetSwitch(name, out value) ? value.ToString() : "unspecified";
        }
        internal static string ExceptionSummary(Exception error)
        {
            var text = new StringBuilder();
            for (int depth = 0; error != null && depth < 8; depth++, error = error.InnerException)
            {
                if (depth > 0) text.Append(" <- ");
                text.Append(error.GetType().FullName).Append(" hresult=0x").Append(error.HResult.ToString("X8"));
                var web = error as WebException;
                if (web != null) text.Append(" web_status=").Append(web.Status);
                var socket = error as SocketException;
                if (socket != null) text.Append(" socket=").Append(socket.SocketErrorCode).Append(" native=").Append(socket.NativeErrorCode);
                var frames = new System.Diagnostics.StackTrace(error, false).GetFrames();
                if (frames != null)
                    for (int index = 0; index < Math.Min(frames.Length, 6); index++)
                    {
                        var method = frames[index].GetMethod();
                        if (method != null) text.Append(" at=").Append(method.DeclaringType == null ? "" : method.DeclaringType.FullName).Append('.').Append(method.Name);
                    }
            }
            return text.ToString();
        }
        internal static string NetworkHint(Exception error)
        {
            var web = error as WebException;
            if (web == null) return "网络或客户端处理失败";
            switch (web.Status)
            {
                case WebExceptionStatus.NameResolutionFailure: return "服务器域名解析失败";
                case WebExceptionStatus.ProxyNameResolutionFailure: return "代理服务器域名解析失败";
                case WebExceptionStatus.ConnectFailure: return "无法建立服务器或代理连接";
                case WebExceptionStatus.TrustFailure: return "HTTPS 证书信任校验失败";
                case WebExceptionStatus.SecureChannelFailure: return "TLS 安全通道协商失败";
                case WebExceptionStatus.Timeout: return "网络请求超时";
                case WebExceptionStatus.SendFailure: return "请求发送失败";
                case WebExceptionStatus.ReceiveFailure: return "响应接收失败";
                case WebExceptionStatus.ConnectionClosed: return "连接被服务器或中间网络设备关闭";
                default: return "网络请求失败（" + web.Status + "）";
            }
        }
        internal static string HttpHint(int status)
        {
            switch (status)
            {
                case 400: return "请求参数或工作流输入校验失败";
                case 401: return "API 密钥未通过鉴权";
                case 403: return "接口权限或访问策略拒绝请求";
                case 404: return "API 地址或应用不存在";
                case 407: return "代理服务器要求身份验证";
                case 408: case 504: return "服务器或网关等待超时";
                case 413: return "请求内容超过服务器限制";
                case 429: return "请求限流或额度不足";
                default: return status >= 300 && status < 400 ? "服务器返回重定向；未跟随跳转" : status >= 500 ? "服务器或网关内部错误" : "接口拒绝请求";
            }
        }
        internal static string ClassifyServerFields(JToken code, JToken error, JToken message)
        {
            foreach (var value in new[] { code, error, message })
            {
                string category = ServerCategory(value);
                if (category != "unclassified" && category != "not_provided") return category;
            }
            return "unclassified";
        }
        internal static string ServerCategory(JToken value)
        {
            // Match known categories, never echo a free-text server error/code.
            string text = value == null ? "" : value.ToString().ToLowerInvariant();
            if (text.Contains("unauthorized") || text.Contains("invalid_api_key") || text.Contains("invalid api key")) return "authentication";
            if (text.Contains("quota") || text.Contains("credit") || text.Contains("balance")) return "quota";
            if (text.Contains("rate_limit") || text.Contains("rate limit")) return "rate_limit";
            if (text.Contains("timeout") || text.Contains("timed out")) return "timeout";
            if (text.Contains("required") || text.Contains("invalid_param") || text.Contains("validation")) return "input_validation";
            if (text.Contains("model") || text.Contains("provider")) return "model_provider";
            return string.IsNullOrWhiteSpace(text) ? "not_provided" : "unclassified";
        }
        internal static string ServerHint(string category)
        {
            switch (category)
            {
                case "authentication": return "服务器报告鉴权失败";
                case "quota": return "服务器报告额度不足";
                case "rate_limit": return "服务器报告请求限流";
                case "timeout": return "服务器报告执行超时";
                case "input_validation": return "服务器报告输入参数校验失败";
                case "model_provider": return "服务器报告模型或供应商调用异常";
                default: return "工作流执行失败或尚未完成";
            }
        }
    }
}
