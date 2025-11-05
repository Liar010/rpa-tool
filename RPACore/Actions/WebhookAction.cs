using System;
using System.Net.Http;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace RPACore.Actions;

/// <summary>
/// Webhookサービス種類
/// </summary>
public enum WebhookServiceType
{
    Custom,      // カスタム（自由にJSON編集）
    Discord,     // Discord Webhook
    Slack,       // Slack Incoming Webhook
    Teams,       // Microsoft Teams Incoming Webhook
    GoogleChat   // Google Chat Webhook
}

/// <summary>
/// Webhook通知を送信するアクション
/// </summary>
public class WebhookAction : ActionBase
{
    /// <summary>Webhook URL</summary>
    public string WebhookUrl { get; set; } = string.Empty;

    /// <summary>サービス種類</summary>
    public WebhookServiceType ServiceType { get; set; } = WebhookServiceType.GoogleChat;

    /// <summary>送信メッセージ</summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>カスタムペイロード（JSON形式、ServiceType=Customの場合のみ使用）</summary>
    public string CustomPayload { get; set; } = string.Empty;

    /// <summary>タイムアウト（ミリ秒）</summary>
    public int TimeoutMs { get; set; } = 10000; // デフォルト10秒

    public override string Name => "Webhook通知";

    public override string Description
    {
        get
        {
            var serviceName = ServiceType switch
            {
                WebhookServiceType.Discord => "Discord",
                WebhookServiceType.Slack => "Slack",
                WebhookServiceType.Teams => "Teams",
                WebhookServiceType.GoogleChat => "Google Chat",
                _ => "カスタム"
            };

            var messagePreview = Message.Length > 30
                ? Message.Substring(0, 30) + "..."
                : Message;

            return $"Webhook通知 ({serviceName}): {messagePreview}";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(WebhookUrl))
            {
                LogError("Webhook URLが指定されていません");
                return false;
            }

            LogInfo($"Webhook送信開始: {ServiceType}");
            LogDebug($"URL: {WebhookUrl}");

            using var httpClient = new HttpClient();
            httpClient.Timeout = TimeSpan.FromMilliseconds(TimeoutMs);

            HttpResponseMessage? response = null;

            switch (ServiceType)
            {
                case WebhookServiceType.Discord:
                    response = await SendDiscordWebhook(httpClient);
                    break;

                case WebhookServiceType.Slack:
                    response = await SendSlackWebhook(httpClient);
                    break;

                case WebhookServiceType.Teams:
                    response = await SendTeamsWebhook(httpClient);
                    break;

                case WebhookServiceType.GoogleChat:
                    response = await SendGoogleChatWebhook(httpClient);
                    break;

                case WebhookServiceType.Custom:
                    response = await SendCustomWebhook(httpClient);
                    break;
            }

            if (response != null && response.IsSuccessStatusCode)
            {
                LogInfo($"Webhook送信成功: {(int)response.StatusCode} {response.ReasonPhrase}");
                return true;
            }
            else if (response != null)
            {
                var responseBody = await response.Content.ReadAsStringAsync();
                LogError($"Webhook送信失敗: {(int)response.StatusCode} {response.ReasonPhrase}");
                LogDebug($"レスポンス: {responseBody}");
                return false;
            }
            else
            {
                LogError("Webhook送信失敗: レスポンスがnullです");
                return false;
            }
        }
        catch (HttpRequestException ex)
        {
            LogError($"HTTP通信エラー: {ex.Message}");
            return false;
        }
        catch (TaskCanceledException ex)
        {
            LogError($"タイムアウト: {ex.Message}");
            return false;
        }
        catch (Exception ex)
        {
            LogError($"予期しないエラー: {ex.Message}");
            return false;
        }
    }

    private async Task<HttpResponseMessage> SendDiscordWebhook(HttpClient httpClient)
    {
        var payload = $"{{\"content\":\"{EscapeJson(Message)}\"}}";
        LogDebug($"Discord ペイロード: {payload}");

        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        return await httpClient.PostAsync(WebhookUrl, content);
    }

    private async Task<HttpResponseMessage> SendSlackWebhook(HttpClient httpClient)
    {
        var payload = $"{{\"text\":\"{EscapeJson(Message)}\"}}";
        LogDebug($"Slack ペイロード: {payload}");

        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        return await httpClient.PostAsync(WebhookUrl, content);
    }

    private async Task<HttpResponseMessage> SendTeamsWebhook(HttpClient httpClient)
    {
        var payload = $"{{\"text\":\"{EscapeJson(Message)}\"}}";
        LogDebug($"Teams ペイロード: {payload}");

        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        return await httpClient.PostAsync(WebhookUrl, content);
    }

    private async Task<HttpResponseMessage> SendGoogleChatWebhook(HttpClient httpClient)
    {
        var payload = $"{{\"text\":\"{EscapeJson(Message)}\"}}";
        LogDebug($"Google Chat ペイロード: {payload}");

        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        return await httpClient.PostAsync(WebhookUrl, content);
    }

    private async Task<HttpResponseMessage> SendCustomWebhook(HttpClient httpClient)
    {
        var payload = string.IsNullOrWhiteSpace(CustomPayload)
            ? $"{{\"message\":\"{EscapeJson(Message)}\"}}"
            : CustomPayload;

        LogDebug($"カスタム ペイロード: {payload}");

        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        return await httpClient.PostAsync(WebhookUrl, content);
    }

    /// <summary>
    /// JSON文字列エスケープ
    /// </summary>
    private string EscapeJson(string text)
    {
        return text
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");
    }
}
