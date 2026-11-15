using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Outline.Addin;

internal static class LiveTests
{
    internal static async Task Run(string keyFile)
    {
        if (string.IsNullOrWhiteSpace(keyFile) || new FileInfo(keyFile).Length > 65536) throw new InvalidDataException();
        var source = JObject.Parse(File.ReadAllText(keyFile));
        var profile = new GkdAiProfile
        {
            Name = "隔离 Dify 实测", Provider = "dify", Endpoint = "https://aibuilder.hdec.com/v1/workflows/run",
            Model = "hdec235", PromptInput = "question", RouteInput = "route", AnswerOutput = "answer",
            ApiKey = (string)source["token"], TimeoutSeconds = 120
        };
        using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(120)))
        {
            var textOnly = await GkdAiClient.ExecuteAsync(profile, "仅回复 TEXT_ONLY_OK，不要解释。", null, "单值", timeout.Token);
            if (!Equals("TEXT_ONLY_OK", textOnly.Value)) throw new InvalidDataException();
            Console.WriteLine("LIVE PASS text-only no-data request elapsed_ms=" + textOnly.ElapsedMilliseconds);
            var scalar = await GkdAiClient.ExecuteAsync(profile, "将输入二维数据中的两个数字相加。仅返回数字42。", new object[,] { { 40, 2 } }, "单值", timeout.Token);
            if (!Equals("42", scalar.Value)) throw new InvalidDataException();
            Console.WriteLine("LIVE PASS scalar value=42 elapsed_ms=" + scalar.ElapsedMilliseconds);
            var reply = await GkdAiClient.ExecuteAsync(profile,
                "将输入数据原样制表，title必须为测试表，表头为名称和值；不要增减数据或添加合计。输出严格JSON。",
                new object[,] { { "名称", "值" }, { "甲", 12 }, { "乙", 30 } }, "表格", timeout.Token);
            var table = reply.Value as object[,];
            if (table == null || table.GetLength(0) != 5 || table.GetLength(1) != 2 || !Equals("测试表", table[0, 0]) || !Equals("", table[0, 1])
                || !Equals("", table[1, 0]) || !Equals("", table[1, 1]) || !Equals("名称", table[2, 0]) || !Equals("值", table[2, 1])
                || !Equals("甲", table[3, 0]) || !Equals("乙", table[4, 0]) || Convert.ToDouble(table[3, 1]) != 12 || Convert.ToDouble(table[4, 1]) != 30)
                throw new InvalidDataException();
            Console.WriteLine("LIVE PASS titled table rows=5 cols=2 title/blank/header/data verified elapsed_ms=" + reply.ElapsedMilliseconds);
        }
    }
}
