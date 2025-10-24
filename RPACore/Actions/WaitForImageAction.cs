using System.Drawing;
using System.Text.Json.Serialization;

namespace RPACore.Actions;

/// <summary>
/// 画像待機モード
/// </summary>
public enum WaitForImageMode
{
    Appear,     // 画像が出現するまで待機
    Disappear   // 画像が消失するまで待機
}

/// <summary>
/// 画像出現/消失待機アクション
///
/// 指定された画像が画面に出現または消失するまで待機します。
/// webedi等のページ読み込み待ちや、処理完了待ちに使用できます。
///
/// 【使用例】
/// - ページ読み込み待ち: ローディング画像が消えるまで待機
/// - ボタン表示待ち: 「次へ」ボタンが表示されるまで待機
/// - 処理完了待ち: 「処理中」表示が消えるまで待機
///
/// 【注意事項】
/// - タイムアウト時間を適切に設定してください（デフォルト: 30秒）
/// - ポーリング間隔が短すぎるとCPU使用率が上がります（デフォルト: 500ms推奨）
/// </summary>
public class WaitForImageAction : ActionBase
{
    /// <summary>
    /// 待機モード（出現 or 消失）
    /// </summary>
    public WaitForImageMode Mode { get; set; } = WaitForImageMode.Appear;

    /// <summary>
    /// テンプレート画像のパス
    /// </summary>
    public string? TemplateImagePath { get; set; }

    /// <summary>
    /// 一致率のしきい値（0.0～1.0）
    /// </summary>
    public double MatchThreshold { get; set; } = 0.8;

    /// <summary>
    /// タイムアウト時間（ミリ秒）
    /// </summary>
    public int TimeoutMs { get; set; } = 30000; // デフォルト30秒

    /// <summary>
    /// ポーリング間隔（ミリ秒）
    /// </summary>
    public int PollingIntervalMs { get; set; } = 500; // デフォルト500ms

    /// <summary>
    /// マルチスケールマッチングを使用するか
    /// </summary>
    public bool UseMultiScale { get; set; } = true;

    /// <summary>
    /// ウィンドウ参照方法
    /// </summary>
    public MouseAction.WindowReferenceMethod WindowReference { get; set; } = MouseAction.WindowReferenceMethod.None;

    /// <summary>
    /// ターゲットウィンドウタイトル（WindowReference = WindowTitle の場合）
    /// </summary>
    public string? TargetWindowTitle { get; set; }

    /// <summary>
    /// ターゲット起動アクション番号（WindowReference = LaunchActionIndex の場合）
    /// </summary>
    public int TargetLaunchActionIndex { get; set; }

    /// <summary>
    /// 検索範囲 X座標（0の場合は全画面）
    /// </summary>
    public int SearchAreaX { get; set; }

    /// <summary>
    /// 検索範囲 Y座標（0の場合は全画面）
    /// </summary>
    public int SearchAreaY { get; set; }

    /// <summary>
    /// 検索範囲 幅（0の場合は全画面）
    /// </summary>
    public int SearchAreaWidth { get; set; }

    /// <summary>
    /// 検索範囲 高さ（0の場合は全画面）
    /// </summary>
    public int SearchAreaHeight { get; set; }

    [JsonIgnore]
    public override string Name => "画像待機";

    [JsonIgnore]
    public override string Description
    {
        get
        {
            string modeText = Mode == WaitForImageMode.Appear ? "出現" : "消失";
            string imageName = System.IO.Path.GetFileName(TemplateImagePath ?? "未設定");
            return $"画像{modeText}待機: {imageName} (タイムアウト: {TimeoutMs}ms)";
        }
    }

    public override bool Validate()
    {
        if (string.IsNullOrEmpty(TemplateImagePath))
        {
            LastError = "テンプレート画像が指定されていません";
            return false;
        }

        if (!System.IO.File.Exists(TemplateImagePath))
        {
            LastError = $"テンプレート画像が見つかりません: {TemplateImagePath}";
            return false;
        }

        if (MatchThreshold < 0 || MatchThreshold > 1)
        {
            LastError = "一致率は0.0～1.0の範囲で指定してください";
            return false;
        }

        if (TimeoutMs <= 0)
        {
            LastError = "タイムアウト時間は正の値で指定してください";
            return false;
        }

        if (PollingIntervalMs <= 0)
        {
            LastError = "ポーリング間隔は正の値で指定してください";
            return false;
        }

        return true;
    }

    public override async Task<bool> ExecuteAsync()
    {
        LogInfo($"画像{(Mode == WaitForImageMode.Appear ? "出現" : "消失")}待機を開始: {TemplateImagePath}");
        LogInfo($"タイムアウト: {TimeoutMs}ms, ポーリング間隔: {PollingIntervalMs}ms");

        // 検索範囲の決定
        Rectangle searchArea;
        if (WindowReference != MouseAction.WindowReferenceMethod.None)
        {
            // ウィンドウ参照の場合
            IntPtr targetHwnd = IntPtr.Zero;

            if (WindowReference == MouseAction.WindowReferenceMethod.WindowTitle)
            {
                if (string.IsNullOrEmpty(TargetWindowTitle))
                {
                    LastError = "ターゲットウィンドウタイトルが指定されていません";
                    LogError(LastError);
                    return false;
                }

                targetHwnd = WindowHelper.FindWindowByTitle(TargetWindowTitle);
                if (targetHwnd == IntPtr.Zero)
                {
                    LastError = $"ウィンドウが見つかりません: {TargetWindowTitle}";
                    LogError(LastError);
                    return false;
                }
                LogInfo($"ウィンドウを検索範囲に設定: {TargetWindowTitle}");
            }
            else if (WindowReference == MouseAction.WindowReferenceMethod.LaunchActionIndex)
            {
                if (Context == null)
                {
                    LastError = "実行コンテキストが設定されていません";
                    LogError(LastError);
                    return false;
                }

                targetHwnd = WindowHelper.FindWindowByLaunchAction(Context, TargetLaunchActionIndex);
                if (targetHwnd == IntPtr.Zero)
                {
                    LastError = $"起動アクション #{TargetLaunchActionIndex} のウィンドウが見つかりません";
                    LogError(LastError);
                    return false;
                }
                LogInfo($"起動アクション #{TargetLaunchActionIndex} のウィンドウを検索範囲に設定");
            }

            searchArea = WindowHelper.GetWindowRectangle(targetHwnd);

            // ウィンドウをアクティブ化
            WindowHelper.ActivateWindow(targetHwnd);
            await Task.Delay(300); // アクティブ化待ち
        }
        else if (SearchAreaWidth > 0 && SearchAreaHeight > 0)
        {
            // 手動指定の検索範囲
            searchArea = new Rectangle(SearchAreaX, SearchAreaY, SearchAreaWidth, SearchAreaHeight);
            LogInfo($"検索範囲: ({SearchAreaX}, {SearchAreaY}) - {SearchAreaWidth}x{SearchAreaHeight}");
        }
        else
        {
            // 全画面
            searchArea = Rectangle.Empty;
            LogInfo("検索範囲: 全画面");
        }

        // タイムアウト用のストップウォッチ
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        int attemptCount = 0;

        while (stopwatch.ElapsedMilliseconds < TimeoutMs)
        {
            attemptCount++;

            try
            {
                // 画像を検索
                using var imageService = new ImageMatchingService();
                var result = imageService.FindImage(
                    TemplateImagePath!,
                    MatchThreshold,
                    UseMultiScale,
                    searchArea
                );

                bool imageFound = result != null;

                // モードに応じて判定
                if (Mode == WaitForImageMode.Appear && imageFound)
                {
                    LogInfo($"画像が出現しました: ({result!.Value.X}, {result.Value.Y}), 試行回数: {attemptCount}");
                    return true;
                }
                else if (Mode == WaitForImageMode.Disappear && !imageFound)
                {
                    LogInfo($"画像が消失しました, 試行回数: {attemptCount}");
                    return true;
                }

                // まだ条件を満たしていない
                LogDebug($"試行 {attemptCount}: 画像{(Mode == WaitForImageMode.Appear ? "未出現" : "まだ存在")}");
            }
            catch (Exception ex)
            {
                LogError($"画像検索エラー (試行 {attemptCount}): {ex.Message}");
            }

            // ポーリング間隔待機
            await Task.Delay(PollingIntervalMs);
        }

        // タイムアウト
        LastError = $"タイムアウト: 画像が{(Mode == WaitForImageMode.Appear ? "出現しませんでした" : "消失しませんでした")} ({stopwatch.ElapsedMilliseconds}ms, {attemptCount}回試行)";
        LogError(LastError);
        return false;
    }
}
