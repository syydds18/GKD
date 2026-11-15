using System;
using System.Collections.Generic;

namespace Outline.Addin
{
    internal sealed class GkdAiProfile
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("D");
        public string Name { get; set; } = "新配置";
        public string Provider { get; set; } = "dify";
        public string Endpoint { get; set; } = "https://aibuilder.hdec.com/v1/workflows/run";
        public string Model { get; set; } = "";
        public string ApiKey { get; set; } = "";
        public string PromptInput { get; set; } = "question";
        public string RouteInput { get; set; } = "route";
        public string AnswerOutput { get; set; } = "answer";
        public int TimeoutSeconds { get; set; } = 120;

        public GkdAiProfile Clone() { return (GkdAiProfile)MemberwiseClone(); }

        public void Validate() { ValidateConfiguration(true); }

        // Empty keys are allowed only while saving/editing a placeholder configuration.
        internal void ValidateConfiguration(bool requireKey)
        {
            Guid id;
            if (!Guid.TryParseExact(Id, "D", out id) || id == Guid.Empty) throw new ArgumentException("配置标识无效。");
            if (string.IsNullOrWhiteSpace(Name) || Name.Length > 100) throw new ArgumentException("配置名称不能为空且不能超过100字。");
            if (Provider != "openai" && Provider != "dify") throw new ArgumentException("服务类型必须为 openai 或 dify。");
            Uri uri;
            if (string.IsNullOrWhiteSpace(Endpoint) || Endpoint.Length > 2048 || HasControl(Endpoint) || Endpoint != Endpoint.Trim() || !Uri.TryCreate(Endpoint, UriKind.Absolute, out uri)
                || (uri.Scheme != Uri.UriSchemeHttps && !(uri.Scheme == Uri.UriSchemeHttp && uri.IsLoopback))
                || !string.IsNullOrEmpty(uri.UserInfo) || !string.IsNullOrEmpty(uri.Query) || !string.IsNullOrEmpty(uri.Fragment))
                throw new ArgumentException("接口地址必须为完整 HTTPS 地址（本机调试可用 HTTP），且不能含凭据、查询参数或片段。");
            if ((requireKey && string.IsNullOrWhiteSpace(ApiKey)) || (ApiKey != null && (ApiKey.Length > 8192 || HasControl(ApiKey))))
                throw new ArgumentException("请填写有效的 API 密钥。");
            if (Model != null && (Model.Length > 256 || HasControl(Model)))
                throw new ArgumentException("模型名称或路由无效；可留空，由服务端决定默认值。");
            if (TimeoutSeconds < 1 || TimeoutSeconds > 600) throw new ArgumentException("超时必须为1至600秒。");
            if (Provider == "dify")
            {
                if (!ValidField(PromptInput) || !ValidField(AnswerOutput) || (!string.IsNullOrEmpty(RouteInput) && !ValidField(RouteInput)))
                    throw new ArgumentException("工作流输入输出变量名无效。");
                if (!string.IsNullOrWhiteSpace(Model) && !string.IsNullOrEmpty(RouteInput) && PromptInput == RouteInput)
                    throw new ArgumentException("问题与路由输入不能使用同一变量名。");
            }
        }
        private static bool HasControl(string value)
        {
            foreach (char ch in value) if (char.IsControl(ch)) return true;
            return false;
        }
        private static bool ValidField(string value) { return !string.IsNullOrWhiteSpace(value) && value.Length <= 128 && !HasControl(value); }
    }

    internal sealed class GkdAiSettings
    {
        public List<GkdAiProfile> Profiles { get; set; } = new List<GkdAiProfile>();
        public string DefaultProfileId { get; set; }
    }
}
