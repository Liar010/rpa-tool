namespace RPACore.Actions;

/// <summary>
/// 待機処理の種類
/// </summary>
public enum WaitType
{
    FixedTime,      // 固定時間待機
    WindowExists    // ウィンドウ出現待機
}

/// <summary>
/// 待機処理アクション
/// </summary>
public class WaitAction : ActionBase
{
    public WaitType WaitType { get; set; } = WaitType.FixedTime;
    public int Milliseconds { get; set; } = 1000;
    public string WindowTitle { get; set; } = string.Empty;
    public int TimeoutSeconds { get; set; } = 30;

    public WaitAction()
    {
        Name = "待機";
        Description = "指定時間待機します";
    }

    public WaitAction(int milliseconds)
    {
        Name = "待機";
        Description = $"{milliseconds}ミリ秒待機";
        WaitType = WaitType.FixedTime;
        Milliseconds = milliseconds;
    }

    public WaitAction(string windowTitle, int timeoutSeconds = 30)
    {
        Name = "ウィンドウ待機";
        Description = $"ウィンドウ \"{windowTitle}\" の出現を待機";
        WaitType = WaitType.WindowExists;
        WindowTitle = windowTitle;
        TimeoutSeconds = timeoutSeconds;
    }

    public override bool Validate()
    {
        if (WaitType == WaitType.FixedTime && Milliseconds <= 0)
        {
            LogError("待機時間が0以下です");
            return false;
        }

        if (WaitType == WaitType.WindowExists && string.IsNullOrEmpty(WindowTitle))
        {
            LogError("ウィンドウタイトルが空です");
            return false;
        }

        return true;
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            if (!Validate())
                return false;

            switch (WaitType)
            {
                case WaitType.FixedTime:
                    await WaitFixedTimeAsync();
                    break;
                case WaitType.WindowExists:
                    return await WaitForWindowAsync();
            }

            return true;
        }
        catch (Exception ex)
        {
            LogError($"待機処理エラー: {ex.Message}");
            return false;
        }
    }

    private async Task WaitFixedTimeAsync()
    {
        LogInfo($"{Milliseconds}ミリ秒待機中...");
        await Task.Delay(Milliseconds);
    }

    private async Task<bool> WaitForWindowAsync()
    {
        LogInfo($"ウィンドウ \"{WindowTitle}\" の出現を待機中...");

        var startTime = DateTime.Now;
        while ((DateTime.Now - startTime).TotalSeconds < TimeoutSeconds)
        {
            IntPtr hwnd = NativeMethods.FindWindow(null, WindowTitle);
            if (hwnd != IntPtr.Zero)
            {
                LogInfo($"ウィンドウが見つかりました: {WindowTitle}");
                return true;
            }

            await Task.Delay(500); // 0.5秒ごとにチェック
        }

        LogError($"タイムアウト: ウィンドウ \"{WindowTitle}\" が見つかりませんでした");
        return false;
    }
}
