using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Outline.Addin
{
    internal sealed class GkdAiReply
    {
        public object Value { get; set; }
        public long ElapsedMilliseconds { get; set; }
    }

    internal static class GkdAiClient
    {
        private const int MaxResponseBytes = 1048576;
        public static async Task<GkdAiReply> ExecuteAsync(GkdAiProfile profile, string prompt, object[,] data, string mode, CancellationToken token)
        {
            var watch = Stopwatch.StartNew();
            string id = Guid.NewGuid().ToString("N").Substring(0, 12);
            string stage = "配置校验", hint = "配置或输入无效";
            int httpStatus = 0;
            HttpWebRequest request = null;
            CancellationTokenSource timeout = null;
            try
            {
                if (profile == null) throw new ArgumentException("请选择 AI 配置。");
                profile = profile.Clone();
                profile.Validate();
                stage = "构造请求";
                string question = GkdAiProtocol.BuildPrompt(prompt, data, mode);
                JObject body;
                if (profile.Provider == "openai")
                {
                    body = new JObject
                    {
                        ["stream"] = false,
                        ["messages"] = new JArray(new JObject { ["role"] = "user", ["content"] = question })
                    };
                    if (!string.IsNullOrWhiteSpace(profile.Model)) body["model"] = profile.Model;
                }
                else
                {
                    var inputs = new JObject { [profile.PromptInput] = question };
                    if (!string.IsNullOrEmpty(profile.RouteInput) && !string.IsNullOrWhiteSpace(profile.Model)) inputs[profile.RouteInput] = profile.Model;
                    body = new JObject { ["inputs"] = inputs, ["response_mode"] = "blocking", ["user"] = "gkd-excel" };
                }
                timeout = new CancellationTokenSource(TimeSpan.FromSeconds(profile.TimeoutSeconds));
                using (var linked = CancellationTokenSource.CreateLinkedTokenSource(token, timeout.Token))
                {
                    token.ThrowIfCancellationRequested();
                    // Keep existing system protocols and add TLS 1.2 for .NET Framework hosts
                    // whose defaults still expose only SSL3/TLS1.0. Never weaken validation.
                    ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                    request = (HttpWebRequest)WebRequest.Create(profile.Endpoint);
                    request.Method = "POST";
                    request.AllowAutoRedirect = false;
                    request.AllowWriteStreamBuffering = false;
                    request.AutomaticDecompression = DecompressionMethods.None;
                    request.ContentType = "application/json; charset=utf-8";
                    request.Accept = "application/json";
                    request.Headers[HttpRequestHeader.Authorization] = "Bearer " + profile.ApiKey;
                    request.Timeout = profile.TimeoutSeconds * 1000;
                    request.ReadWriteTimeout = profile.TimeoutSeconds * 1000;
                    byte[] payload = Encoding.UTF8.GetBytes(body.ToString(Formatting.None));
                    request.ContentLength = payload.Length;
                    GkdAiDiagnostics.Write(id, stage, GkdAiDiagnostics.EnvironmentSummary(request) + " provider=" + profile.Provider + " request_bytes=" + payload.Length);
                    using (linked.Token.Register(request.Abort))
                    {
                        stage = "建立连接并发送请求";
                        hint = "连接、TLS 或请求发送失败";
                        GkdAiDiagnostics.Write(id, stage, "begin");
                        using (var output = await request.GetRequestStreamAsync().ConfigureAwait(false))
                        {
                            stage = "写入请求";
                            await output.WriteAsync(payload, 0, payload.Length, linked.Token).ConfigureAwait(false);
                        }
                        stage = "等待 HTTP 响应";
                        hint = "未收到有效 HTTP 响应";
                        GkdAiDiagnostics.Write(id, stage, "request_sent=true elapsed_ms=" + watch.ElapsedMilliseconds);
                        string responseText;
                        using (var response = (HttpWebResponse)await request.GetResponseAsync().ConfigureAwait(false))
                        {
                            httpStatus = (int)response.StatusCode;
                            GkdAiDiagnostics.Write(id, "HTTP 响应", "http_status=" + httpStatus + " content_length=" + response.ContentLength);
                            hint = GkdAiDiagnostics.HttpHint(httpStatus);
                            if (httpStatus < 200 || httpStatus >= 300) throw new InvalidDataException();
                            stage = "读取响应";
                            hint = "响应超过大小限制或读取失败";
                            responseText = await ReadResponse(response, linked.Token).ConfigureAwait(false);
                        }
                        linked.Token.ThrowIfCancellationRequested();
                        stage = "解析响应 JSON";
                        hint = "响应不是有效 JSON；可能返回了登录页或网关页面";
                        var root = GkdAiProtocol.ReadJson(responseText) as JObject;
                        if (root == null) throw new InvalidDataException();
                        JToken answer;
                        if (profile.Provider == "openai")
                        {
                            stage = "检查模型响应";
                            hint = "模型未完整返回答案";
                            var choices = root["choices"] as JArray;
                            var first = choices != null && choices.Count > 0 ? choices[0] as JObject : null;
                            var message = first == null ? null : first["message"] as JObject;
                            string reason = first == null ? null : (string)first["finish_reason"];
                            if (reason != null && reason != "stop") throw new InvalidDataException();
                            answer = message == null ? null : message["content"];
                        }
                        else
                        {
                            stage = "检查工作流状态";
                            var workflow = root["data"] as JObject;
                            bool succeeded = workflow != null && workflow["status"] != null && workflow["status"].Type == JTokenType.String && (string)workflow["status"] == "succeeded";
                            string category = GkdAiDiagnostics.ServerCategory(workflow == null ? root["message"] : workflow["error"]);
                            GkdAiDiagnostics.Write(id, stage, "workflow_succeeded=" + succeeded + " error_category=" + category);
                            hint = GkdAiDiagnostics.ServerHint(category);
                            if (!succeeded) throw new InvalidDataException();
                            var outputs = workflow["outputs"] as JObject;
                            answer = outputs == null ? null : outputs[profile.AnswerOutput];
                        }
                        stage = "提取答案字段";
                        hint = "配置的答案字段缺失或不是文本；请核对输出映射";
                        if (answer == null || answer.Type != JTokenType.String) throw new InvalidDataException();
                        stage = "校验单值或表格格式";
                        hint = "答案为空、超限，或表格 JSON 的标题、列数和数据格式不符合要求";
                        object value = GkdAiProtocol.ParseAnswer((string)answer, mode);
                        token.ThrowIfCancellationRequested();
                        GkdAiDiagnostics.Write(id, "完成", "http_status=" + httpStatus + " elapsed_ms=" + watch.ElapsedMilliseconds);
                        return new GkdAiReply { Value = value, ElapsedMilliseconds = watch.ElapsedMilliseconds };
                    }
                }
            }
            catch (Exception error) when (!(error is OutOfMemoryException) && !(error is StackOverflowException) && !(error is AccessViolationException))
            {
                // Capture the original request outcome before bounded diagnostic collection;
                // reading an error body must not turn a 401/500 into a timeout.
                bool cancelledBeforeDiagnostics = token.IsCancellationRequested;
                bool timedOutBeforeDiagnostics = timeout != null && timeout.IsCancellationRequested;
                var web = error as WebException;
                if (web != null)
                {
                    var response = web.Response as HttpWebResponse;
                    if (response != null)
                    {
                        httpStatus = (int)response.StatusCode;
                        hint = GkdAiDiagnostics.HttpHint(httpStatus);
                        // Error bodies may echo confidential data. Log only a mapped category,
                        // and bound reading with a separate deadline after the request deadline.
                        using (response)
                        using (var readTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(2)))
                        using (var readLinked = CancellationTokenSource.CreateLinkedTokenSource(token, readTimeout.Token))
                        using (readLinked.Token.Register(request.Abort))
                        {
                            try
                            {
                                var root = JObject.Parse(await ReadResponse(response, readLinked.Token).ConfigureAwait(false));
                                string category = GkdAiDiagnostics.ClassifyServerFields(root["code"], root["error"], root["message"]);
                                GkdAiDiagnostics.Write(id, "HTTP 错误响应", "http_status=" + httpStatus + " error_category=" + category);
                            }
                            catch (Exception) { GkdAiDiagnostics.Write(id, "HTTP 错误响应", "http_status=" + httpStatus + " body_details=omitted"); }
                        }
                    }
                    else hint = GkdAiDiagnostics.NetworkHint(error);
                }
                bool cancelled = cancelledBeforeDiagnostics || token.IsCancellationRequested;
                bool timedOut = timedOutBeforeDiagnostics;
                if (cancelled) hint = "请求已取消";
                else if (timedOut) hint = "请求超过配置的超时时间";
                string path = GkdAiDiagnostics.Write(id, stage, "FAILED http_status=" + httpStatus + " elapsed_ms=" + watch.ElapsedMilliseconds + " cancelled=" + cancelled + " timeout=" + timedOut + " " + GkdAiDiagnostics.ExceptionSummary(error));
                string details = "AI 请求失败：" + stage + "；" + hint + (httpStatus == 0 ? "" : "（HTTP " + httpStatus + "）") + "。编号 " + id + "；" + (path == null ? "日志写入失败，请检查目录权限。" : "详情见诊断日志。");
                if (cancelled) throw new OperationCanceledException("AI 请求已取消；编号 " + id, token);
                if (timedOut) throw new TimeoutException(details);
                // Never preserve an inner exception: its message may contain request data.
                throw new InvalidDataException(details);
            }
            finally { if (timeout != null) timeout.Dispose(); }
        }

        private static async Task<string> ReadResponse(HttpWebResponse response, CancellationToken token)
        {
            if (response.ContentLength > MaxResponseBytes) throw new InvalidDataException();
            using (var input = response.GetResponseStream())
            using (var buffer = new MemoryStream())
            {
                var block = new byte[8192];
                int count;
                while ((count = await input.ReadAsync(block, 0, block.Length, token).ConfigureAwait(false)) > 0)
                {
                    if (buffer.Length + count > MaxResponseBytes) throw new InvalidDataException();
                    buffer.Write(block, 0, count);
                }
                return new UTF8Encoding(false, true).GetString(buffer.ToArray());
            }
        }
    }
}
