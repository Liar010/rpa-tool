using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.Json.Serialization;

namespace RPACore.Actions;

/// <summary>
/// スクロール方向
/// </summary>
public enum ScrollDirection
{
    Down,   // 下にスクロール
    Up,     // 上にスクロール
    Right,  // 右にスクロール
    Left    // 左にスクロール
}

/// <summary>
/// スクロール検索アクション
///
/// 指定された画像が見つかるまでスクロールを繰り返します。
/// webedi等の長いページで要素を探す場合に便利です。
///
/// 【使用例】
/// - リスト内の特定項目を探す
/// - 長いページで「次へ」ボタンを探す
/// - テーブル内の特定行を探す
///
/// 【注意事項】
/// - 最大スクロール回数を設定して無限ループを防いでください
/// - スクロール後の画面安定待ち時間を適切に設定してください
/// - ウィンドウをアクティブにしてから実行してください
/// </summary>
public class ScrollUntilImageAction : ActionBase
{
    [DllImport("user32.dll")]
    private static extern void mouse_event(uint dwFlags, int dx, int dy, int dwData, UIntPtr dwExtraInfo);

    private const uint MOUSEEVENTF_WHEEL = 0x0800;
    private const uint MOUSEEVENTF_HWHEEL = 0x01000;
    private const int WHEEL_DELTA = 120;

    /// <summary>
    /// テンプレート画像のパス
    /// </summary>
    public string? TemplateImagePath { get; set; }

    /// <summary>
    /// 一致率のしきい値（0.0～1.0）
    /// </summary>
    public double MatchThreshold { get; set; } = 0.8;

    /// <summary>
    /// スクロール方向
    /// </summary>
    public ScrollDirection Direction { get; set; } = ScrollDirection.Down;

    /// <summary>
    /// 1回のスクロール量（ホイールクリック数）
    /// </summary>
    public int ScrollAmount { get; set; } = 3;

    /// <summary>
    /// 最大スクロール回数（無限ループ防止）
    /// </summary>
    public int MaxScrollCount { get; set; } = 20;

    /// <summary>
    /// スクロール後の待機時間（ミリ秒）
    /// ページの描画が完了するまで待つ
    /// </summary>
    public int WaitAfterScrollMs { get; set; } = 500;

    /// <summary>
    /// マルチスケールマッチングを使用するか
    /// </summary>
    public bool UseMultiScale { get; set; } = true;

    /// <summary>
    /// 画像発見時にクリックするか
    /// </summary>
    public bool ClickWhenFound { get; set; } = false;

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

    /// <summary>
    /// 発見した画像の位置（クリック時に使用）
    /// </summary>
    [JsonIgnore]
    public Point? FoundPosition { get; private set; }

    [JsonIgnore]
    public override string Name => "スクロール検索";

    [JsonIgnore]
    public override string Description
    {
        get
        {
            string imageName = System.IO.Path.GetFileName(TemplateImagePath ?? "未設定");
            string directionText = Direction switch
            {
                ScrollDirection.Down => "下",
                ScrollDirection.Up => "上",
                ScrollDirection.Right => "右",
                ScrollDirection.Left => "左",
                _ => "?"
            };
            return $"スクロール検索({directionText}): {imageName} (最大{MaxScrollCount}回)";
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

        if (ScrollAmount <= 0)
        {
            LastError = "スクロール量は正の値で指定してください";
            return false;
        }

        if (MaxScrollCount <= 0)
        {
            LastError = "最大スクロール回数は正の値で指定してください";
            return false;
        }

        return true;
    }

    public override async Task<bool> ExecuteAsync()
    {
        LogInfo($"スクロール検索を開始: {TemplateImagePath}, 方向: {Direction}, 最大回数: {MaxScrollCount}");

        // 検索範囲の決定
        Rectangle searchArea;
        IntPtr targetHwnd = IntPtr.Zero;

        if (WindowReference != MouseAction.WindowReferenceMethod.None)
        {
            // ウィンドウ参照の場合
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

        // スクロール検索ループ
        for (int i = 0; i < MaxScrollCount; i++)
        {
            LogDebug($"スクロール試行 {i + 1}/{MaxScrollCount}");

            // ウィンドウが指定されている場合、スクロール前に毎回アクティブ化
            if (targetHwnd != IntPtr.Zero)
            {
                WindowHelper.ActivateWindow(targetHwnd);
                await Task.Delay(100); // アクティブ化待ち

                // カーソルをウィンドウの中央に移動（スクロールを確実に送るため）
                var rect = WindowHelper.GetWindowRectangle(targetHwnd);
                int centerX = rect.Left + rect.Width / 2;
                int centerY = rect.Top + rect.Height / 2;
                NativeMethods.SetCursorPos(centerX, centerY);
                await Task.Delay(50);
            }

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

                if (result != null)
                {
                    LogInfo($"画像を発見しました: ({result.Value.X}, {result.Value.Y}), スクロール回数: {i}");
                    FoundPosition = result.Value;

                    // クリックオプションが有効な場合
                    if (ClickWhenFound)
                    {
                        LogInfo($"発見位置をクリックします: ({result.Value.X}, {result.Value.Y})");
                        NativeMethods.SetCursorPos(result.Value.X, result.Value.Y);
                        await Task.Delay(100);
                        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
                        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                LogError($"画像検索エラー (試行 {i + 1}): {ex.Message}");
            }

            // 最後の試行でなければスクロール
            if (i < MaxScrollCount - 1)
            {
                PerformScroll();
                await Task.Delay(WaitAfterScrollMs);
            }
        }

        // 見つからなかった
        LastError = $"画像が見つかりませんでした: {MaxScrollCount}回スクロールしましたが発見できませんでした";
        LogError(LastError);
        return false;
    }

    /// <summary>
    /// スクロールを実行
    /// </summary>
    private void PerformScroll()
    {
        int wheelDelta = WHEEL_DELTA * ScrollAmount;

        switch (Direction)
        {
            case ScrollDirection.Down:
                mouse_event(MOUSEEVENTF_WHEEL, 0, 0, -wheelDelta, UIntPtr.Zero);
                LogDebug($"下にスクロール: {ScrollAmount}クリック");
                break;

            case ScrollDirection.Up:
                mouse_event(MOUSEEVENTF_WHEEL, 0, 0, wheelDelta, UIntPtr.Zero);
                LogDebug($"上にスクロール: {ScrollAmount}クリック");
                break;

            case ScrollDirection.Right:
                mouse_event(MOUSEEVENTF_HWHEEL, 0, 0, wheelDelta, UIntPtr.Zero);
                LogDebug($"右にスクロール: {ScrollAmount}クリック");
                break;

            case ScrollDirection.Left:
                mouse_event(MOUSEEVENTF_HWHEEL, 0, 0, -wheelDelta, UIntPtr.Zero);
                LogDebug($"左にスクロール: {ScrollAmount}クリック");
                break;
        }
    }
}
