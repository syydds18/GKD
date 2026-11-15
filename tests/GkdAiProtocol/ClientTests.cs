using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Outline.Addin;

internal static class ClientTests
{
    internal static async Task OpenAi()
    {
        using (var server = new Fixture(async ctx =>
        {
            await Task.Delay(60);
            await Fixture.Send(ctx, "{\"choices\":[{\"message\":{\"content\":\"<think>hidden</think>42\"},\"finish_reason\":\"stop\"}]}");
        }))
        {
            var reply = await GkdAiClient.ExecuteAsync(Profile(server), "统计", new object[,] { { 40, 2 } }, null, CancellationToken.None);
            Eq("42", reply.Value); Eq(true, reply.ElapsedMilliseconds >= 40);
            Eq("Bearer fixture-key", server.Authorization); Eq("/v1/chat/completions", server.Path);
            var body = JObject.Parse(server.Body); Eq("fixture-model", (string)body["model"]); Eq(false, (bool)body["stream"]);
            Eq(1, ((JArray)body["messages"]).Count); Eq("user", (string)body["messages"][0]["role"]);
            Eq(true, ((string)body["messages"][0]["content"]).Contains("统计"));
            Eq(null, body["conversation_id"]); Eq(null, body["tools"]);
            Eq(1, server.Count);
        }
        using (var server = new Fixture(ctx => Fixture.Send(ctx, "{\"choices\":[{\"message\":{\"content\":\"ok\"},\"finish_reason\":\"stop\"}]}")))
        {
            var p = Profile(server); p.Model = "";
            Eq("ok", (await GkdAiClient.ExecuteAsync(p, "x", new object[,] { { 1 } }, null, CancellationToken.None)).Value);
            Eq(null, JObject.Parse(server.Body)["model"]);
        }
        using (var server = new Fixture(ctx => Fixture.Send(ctx, "{\"choices\":[{\"message\":{\"content\":\"text-only-ok\"},\"finish_reason\":\"stop\"}]}")))
        {
            var previousProtocols = ServicePointManager.SecurityProtocol;
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                var reply = await GkdAiClient.ExecuteAsync(Profile(server), "只回答问题", null, null, CancellationToken.None);
                Eq("text-only-ok", reply.Value);
                Eq(true, (ServicePointManager.SecurityProtocol & SecurityProtocolType.Tls12) != 0);
            }
            finally { ServicePointManager.SecurityProtocol = previousProtocols; }
            Eq(null, JObject.Parse(server.Body)["data"]);
            Eq(true, ((string)JObject.Parse(server.Body)["messages"][0]["content"]).Contains("\"data\":null"));
        }
    }
    internal static async Task Dify()
    {
        string tableAnswer = new JObject { ["title"] = "测试表", ["data"] = new JArray(new JArray("值"), new JArray(42)) }.ToString(Formatting.None);
        string response = new JObject { ["data"] = new JObject { ["status"] = "succeeded", ["outputs"] = new JObject { ["custom_answer"] = tableAnswer } } }.ToString(Formatting.None);
        using (var server = new Fixture(ctx => Fixture.Send(ctx, response)))
        {
            var p = Profile(server); p.Provider = "dify"; p.Endpoint = server.Url + "v1/workflows/run";
            p.Model = "route-a"; p.PromptInput = "prompt"; p.RouteInput = "model_route"; p.AnswerOutput = "custom_answer";
            var reply = await GkdAiClient.ExecuteAsync(p, "提取", new object[,] { { 42 } }, "表格", CancellationToken.None);
            Eq(42.0, ((object[,])reply.Value)[3, 0]); Eq("测试表", ((object[,])reply.Value)[0, 0]); Eq("", ((object[,])reply.Value)[1, 0]); Eq("值", ((object[,])reply.Value)[2, 0]);
            var body = JObject.Parse(server.Body); Eq("blocking", (string)body["response_mode"]);
            Eq("route-a", (string)body["inputs"]["model_route"]); Eq(true, ((string)body["inputs"]["prompt"]).Contains("提取"));
            Eq(null, body["conversation_id"]); Eq(true, !string.IsNullOrEmpty((string)body["user"]));
        }
        using (var server = new Fixture(ctx => Fixture.Send(ctx, "{\"data\":{\"status\":\"succeeded\",\"outputs\":{\"answer\":\"ok\"}}}")))
        {
            var p = Profile(server); p.Provider = "dify"; p.Model = "";
            Eq("ok", (await GkdAiClient.ExecuteAsync(p, "x", new object[,] { { 1 } }, null, CancellationToken.None)).Value);
            Eq(null, JObject.Parse(server.Body)["inputs"]["route"]);
        }
    }
    internal static async Task Failures()
    {
        foreach (string body in new[] { "{\"data\":{\"status\":\"failed\",\"error\":\"fixture-key PRIVATE-PROMPT\",\"outputs\":{\"answer\":\"misleading\"}}}", "{\"data\":{\"outputs\":{\"answer\":\"misleading\"}}}", "{\"data\":{\"status\":\"succeeded\",\"outputs\":{}}}" })
        using (var server = new Fixture(ctx => Fixture.Send(ctx, body)))
        {
            var p = Profile(server); p.Provider = "dify"; p.Model = "";
            await Fails(() => GkdAiClient.ExecuteAsync(p, "PRIVATE-PROMPT", new object[,] { { 1 } }, null, CancellationToken.None));
        }
        foreach (string body in new[] { "not JSON fixture-key PRIVATE-PROMPT", "{\"choices\":[]}", "{\"choices\":[{\"message\":{\"content\":null}}]}", "{\"choices\":[{\"message\":{\"content\":\"partial\"},\"finish_reason\":\"length\"}]}" })
        using (var server = new Fixture(ctx => Fixture.Send(ctx, body)))
            await Fails(() => GkdAiClient.ExecuteAsync(Profile(server), "PRIVATE-PROMPT", new object[,] { { 1 } }, null, CancellationToken.None));
        using (var server = new Fixture(ctx => { ctx.Response.StatusCode = 401; return Fixture.Send(ctx, "fixture-key PRIVATE-PROMPT"); }))
            await Fails(() => GkdAiClient.ExecuteAsync(Profile(server), "PRIVATE-PROMPT", new object[,] { { 1 } }, null, CancellationToken.None));
    }
    internal static async Task RedirectAndLimits()
    {
        using (var target = new Fixture(ctx => Fixture.Send(ctx, "{}")))
        using (var server = new Fixture(ctx => { ctx.Response.StatusCode = 307; ctx.Response.RedirectLocation = target.Url; return Fixture.Send(ctx, "redirect"); }))
        {
            await Fails(() => GkdAiClient.ExecuteAsync(Profile(server), "PRIVATE-PROMPT", new object[,] { { 1 } }, null, CancellationToken.None));
            Eq(0, target.Count); Eq(1, server.Count);
        }
        foreach (bool chunked in new[] { false, true })
        using (var server = new Fixture(ctx => Fixture.Send(ctx, new string('x', 1048577), chunked)))
            await Fails(() => GkdAiClient.ExecuteAsync(Profile(server), "PRIVATE-PROMPT", new object[,] { { 1 } }, null, CancellationToken.None));
    }
    internal static async Task Cancellation()
    {
        using (var server = new Fixture(async ctx => { await Task.Delay(1500); await Fixture.Send(ctx, "{}"); }))
        using (var cancel = new CancellationTokenSource())
        {
            cancel.CancelAfter(100);
            try { await GkdAiClient.ExecuteAsync(Profile(server), "x", new object[,] { { 1 } }, null, cancel.Token); throw new Exception("Expected cancellation"); }
            catch (OperationCanceledException) { }
        }
        using (var server = new Fixture(async ctx => { await Task.Delay(1500); await Fixture.Send(ctx, "{}"); }))
        {
            var p = Profile(server); p.TimeoutSeconds = 1;
            try { await GkdAiClient.ExecuteAsync(p, "x", new object[,] { { 1 } }, null, CancellationToken.None); throw new Exception("Expected timeout"); }
            catch (TimeoutException e) { Eq(false, e.ToString().Contains("fixture-key")); }
        }
    }
    internal static async Task Diagnostics()
    {
        Eq(true, GkdAiDiagnostics.NetworkHint(new WebException("fixture-key PRIVATE-PROMPT", WebExceptionStatus.TrustFailure)).Contains("证书"));
        Eq(true, GkdAiDiagnostics.NetworkHint(new WebException("private", WebExceptionStatus.SecureChannelFailure)).Contains("TLS"));
        Eq(true, GkdAiDiagnostics.NetworkHint(new WebException("private", WebExceptionStatus.NameResolutionFailure)).Contains("域名"));
        string diagnosticId = null;
        foreach (int status in new[] { 400, 401, 403, 407, 429, 500, 504 })
        using (var server = new Fixture(ctx =>
        {
            ctx.Response.StatusCode = status;
            return Fixture.Send(ctx, "{\"code\":\"invalid_param\",\"message\":\"fixture-key PRIVATE-PROMPT PRIVATE-CELL\"}");
        }))
        {
            try
            {
                await GkdAiClient.ExecuteAsync(Profile(server), "PRIVATE-PROMPT", new object[,] { { "PRIVATE-CELL" } }, null, CancellationToken.None);
                throw new Exception("Expected diagnostic failure");
            }
            catch (InvalidDataException ex)
            {
                Eq(true, ex.Message.Contains("HTTP " + status));
                Eq(true, ex.Message.Contains("编号"));
                Eq(null, ex.InnerException);
                diagnosticId = ex.Message.Split(new[] { "编号 " }, StringSplitOptions.None)[1].Substring(0, 12);
            }
        }
        using (var server = new Fixture(ctx => Fixture.Send(ctx, "PRIVATE-CELL fixture-key PRIVATE-PROMPT")))
        {
            try { await GkdAiClient.ExecuteAsync(Profile(server), "PRIVATE-PROMPT", new object[,] { { "PRIVATE-CELL" } }, null, CancellationToken.None); }
            catch (InvalidDataException ex) { Eq(true, ex.Message.Contains("解析响应 JSON")); }
        }
        string log = File.ReadAllText(GkdAiDiagnostics.LogPath);
        Eq(true, log.Contains("request=" + diagnosticId));
        Eq(true, log.Contains("http_status=401"));
        Eq(true, log.Contains("clr=")); Eq(true, log.Contains("tls_policy="));
        Eq(false, log.Contains("fixture-key")); Eq(false, log.Contains("PRIVATE-PROMPT")); Eq(false, log.Contains("PRIVATE-CELL"));
        Eq(false, GkdAiDiagnostics.ExceptionSummary(new IOException("PRIVATE-CELL", new WebException("fixture-key"))).Contains("fixture-key"));
    }
    private static GkdAiProfile Profile(Fixture server)
    {
        return new GkdAiProfile { Provider = "openai", Endpoint = server.Url + "v1/chat/completions", Model = "fixture-model", ApiKey = "fixture-key" };
    }
    private static async Task Fails(Func<Task<GkdAiReply>> run)
    {
        try { await run(); } catch (InvalidDataException e)
        {
            Eq(false, e.ToString().Contains("fixture-key")); Eq(false, e.ToString().Contains("PRIVATE-PROMPT")); Eq(true, e.Message.Length < 200); return;
        }
        throw new Exception("Expected sanitized failure");
    }
    private static void Eq(object expected, object actual) { if (!Equals(expected, actual)) throw new Exception("Expected [" + expected + "], got [" + actual + "]"); }

    private sealed class Fixture : IDisposable
    {
        private readonly HttpListener listener = new HttpListener();
        private int count;
        public string Url { get; private set; }
        public string Body;
        public string Authorization;
        public string Path;
        public int Count { get { return Volatile.Read(ref count); } }
        internal Fixture(Func<HttpListenerContext, Task> handle)
        {
            var ports = new TcpListener(IPAddress.Loopback, 0); ports.Start(); int port = ((IPEndPoint)ports.LocalEndpoint).Port; ports.Stop();
            Url = "http://127.0.0.1:" + port + "/"; listener.Prefixes.Add(Url); listener.Start();
            Task.Run(async () =>
            {
                while (listener.IsListening)
                {
                    HttpListenerContext ctx;
                    try { ctx = await listener.GetContextAsync(); } catch (Exception) { break; }
                    try
                    {
                        using (var reader = new StreamReader(ctx.Request.InputStream, Encoding.UTF8)) Body = await reader.ReadToEndAsync();
                        Authorization = ctx.Request.Headers["Authorization"]; Path = ctx.Request.Url.AbsolutePath; Interlocked.Increment(ref count);
                        await handle(ctx);
                    }
                    catch (Exception) { }
                    finally { try { ctx.Response.Close(); } catch (Exception) { } }
                }
            });
        }
        internal static async Task Send(HttpListenerContext ctx, string body, bool chunked = false)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(body); ctx.Response.ContentType = "application/json";
            if (chunked) ctx.Response.SendChunked = true; else ctx.Response.ContentLength64 = bytes.Length;
            await ctx.Response.OutputStream.WriteAsync(bytes, 0, bytes.Length);
        }
        public void Dispose() { listener.Close(); }
    }
}
