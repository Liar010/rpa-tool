using System.Drawing;
using System.Text.Json.Serialization;

namespace RPACore.Actions;

/// <summary>
/// マウスクリックの種類
/// </summary>
public enum MouseClickType
{
    LeftClick,      // 左クリック
    RightClick,     // 右クリック
    DoubleClick,    // ダブルクリック
    MiddleClick     // 中央クリック
}

/// <summary>
/// マウス操作アクション
///
/// クリック方法:
/// - 座標指定: X,Y座標を直接指定してクリック
/// - 画像認識: テンプレート画像と一致する位置を検索してクリック
///
/// 【画像認識使用時の注意事項】
/// 1. 対象ウィンドウは最大化またはウィンドウサイズ固定を推奨
///    - ウィンドウサイズが変わると画像サイズも変わり、マッチング精度が低下します
///
/// 2. 解像度・DPI設定の影響
///    - マルチスケールマッチング（推奨）を使用することで、DPI変更や解像度変更に対応できます
///    - 単一スケールの場合、スクリプト作成時と実行時で画面設定が異なると動作しない可能性があります
///
/// 3. テンプレート画像の撮影
///    - クリックしたい要素（ボタンなど）を明確に含むように範囲選択してください
///    - 周囲の余白は最小限にすることで精度が向上します
///    - 背景に動的に変化する要素（アニメーション等）が含まれないようにしてください
///
/// 4. 一致率の調整
///    - デフォルト: 80%
///    - 見つからない場合: 70%程度に下げてください
///    - 誤認識が多い場合: 85-90%に上げてください
///
/// 5. ウィンドウ操作との組み合わせ
///    - 画像認識クリック前に、WindowActionでウィンドウをアクティブ化・最大化することを推奨
///    - 例: #1 ウィンドウ最大化 → #2 画像認識クリック
/// </summary>
public class MouseAction : ActionBase
{
    /// <summary>
    /// クリック方法
    /// </summary>
    public enum ClickMethod
    {
        /// <summary>座標指定</summary>
        Coordinate,
        /// <summary>画像認識</summary>
        ImageMatch
    }

    // === クリック方法 ===

    /// <summary>
    /// クリック方法（デフォルト: 座標指定）
    /// </summary>
    public ClickMethod Method { get; set; } = ClickMethod.Coordinate;

    // === 共通プロパティ ===

    /// <summary>
    /// クリック種類
    /// </summary>
    public MouseClickType ClickType { get; set; } = MouseClickType.LeftClick;

    /// <summary>
    /// クリック後の待機時間（ミリ秒）
    /// </summary>
    public int DelayAfterClick { get; set; } = 100;

    // === 座標指定用プロパティ ===

    /// <summary>
    /// X座標（座標指定時のみ使用）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int X { get; set; }

    /// <summary>
    /// Y座標（座標指定時のみ使用）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int Y { get; set; }

    // === 画像認識用プロパティ ===

    /// <summary>
    /// テンプレート画像のパス（画像認識時のみ使用）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string? TemplateImagePath { get; set; }

    /// <summary>
    /// 一致率の閾値（0.0～1.0、デフォルト: 0.8 = 80%）
    /// 見つからない場合は0.7程度に下げることを推奨
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public double MatchThreshold { get; set; } = 0.8;

    /// <summary>
    /// タイムアウト時間（ミリ秒、デフォルト: 5000ms = 5秒）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int SearchTimeoutMs { get; set; } = 5000;

    /// <summary>
    /// マルチスケールマッチングを使用するか（デフォルト: true）
    /// 解像度やDPI変更に対応するため、通常はtrueを推奨
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool UseMultiScale { get; set; } = true;

    /// <summary>
    /// 検索範囲X座標（0の場合は全画面検索）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int SearchAreaX { get; set; }

    /// <summary>
    /// 検索範囲Y座標（0の場合は全画面検索）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int SearchAreaY { get; set; }

    /// <summary>
    /// 検索範囲の幅（0の場合は全画面検索）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int SearchAreaWidth { get; set; }

    /// <summary>
    /// 検索範囲の高さ（0の場合は全画面検索）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int SearchAreaHeight { get; set; }

    // === ウィンドウ指定（画像認識用） ===

    /// <summary>
    /// ウィンドウ参照方法
    /// </summary>
    public enum WindowReferenceMethod
    {
        /// <summary>指定なし（全画面検索）</summary>
        None,
        /// <summary>ウィンドウタイトルで指定</summary>
        WindowTitle,
        /// <summary>起動アクションの番号で指定</summary>
        LaunchActionIndex
    }

    /// <summary>
    /// ウィンドウ参照方法（画像認識時のみ使用、デフォルト: None = 全画面検索）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public WindowReferenceMethod WindowReference { get; set; } = WindowReferenceMethod.None;

    /// <summary>
    /// 検索対象ウィンドウのタイトル（WindowReference = WindowTitle の場合のみ使用）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string? TargetWindowTitle { get; set; }

    /// <summary>
    /// 起動アクションの番号（WindowReference = LaunchActionIndex の場合のみ使用）
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int TargetLaunchActionIndex { get; set; }

    // === オーバーライド ===

    public override string Name => "マウスクリック";

    public override string Description => Method switch
    {
        ClickMethod.Coordinate => $"{GetClickTypeName(ClickType)}: ({X}, {Y})",
        ClickMethod.ImageMatch => $"{GetClickTypeName(ClickType)}: 画像認識 ({Path.GetFileName(TemplateImagePath) ?? "未設定"})",
        _ => "マウスクリック"
    };

    public override bool Validate()
    {
        switch (Method)
        {
            case ClickMethod.Coordinate:
                if (X < 0 || Y < 0)
                {
                    LastError = "X座標とY座標は0以上である必要があります。";
                    return false;
                }
                break;

            case ClickMethod.ImageMatch:
                if (string.IsNullOrEmpty(TemplateImagePath))
                {
                    LastError = "テンプレート画像のパスが指定されていません。";
                    return false;
                }
                if (!File.Exists(TemplateImagePath))
                {
                    LastError = $"テンプレート画像が見つかりません: {TemplateImagePath}";
                    return false;
                }
                if (MatchThreshold < 0.0 || MatchThreshold > 1.0)
                {
                    LastError = "一致率は0.0～1.0の範囲で指定してください。";
                    return false;
                }
                break;
        }

        return true;
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            if (!Validate())
                return false;

            switch (Method)
            {
                case ClickMethod.Coordinate:
                    return await ExecuteCoordinateClick();

                case ClickMethod.ImageMatch:
                    return await ExecuteImageClick();

                default:
                    LastError = "不明なクリック方法が指定されています。";
                    return false;
            }
        }
        catch (Exception ex)
        {
            LogError($"マウス操作エラー: {ex.Message}");
            LastError = ex.Message;
            return false;
        }
    }

    // === 座標指定クリック ===

    private async Task<bool> ExecuteCoordinateClick()
    {
        LogInfo($"座標 ({X}, {Y}) に移動して{GetClickTypeName(ClickType)}");

        // カーソル移動
        NativeMethods.SetCursorPos(X, Y);
        await Task.Delay(50); // カーソル移動後の安定待ち

        // クリック実行
        await PerformClick(ClickType);

        await Task.Delay(DelayAfterClick);
        return true;
    }

    // === 画像認識クリック ===

    private async Task<bool> ExecuteImageClick()
    {
        LogInfo($"画像認識クリック開始: {Path.GetFileName(TemplateImagePath)}");

        Rectangle? searchArea = null;

        // ウィンドウ指定がある場合、そのウィンドウの矩形を取得
        if (WindowReference != WindowReferenceMethod.None)
        {
            IntPtr hWnd = IntPtr.Zero;

            switch (WindowReference)
            {
                case WindowReferenceMethod.WindowTitle:
                    if (string.IsNullOrEmpty(TargetWindowTitle))
                    {
                        LogError("ウィンドウタイトルが指定されていません");
                        LastError = "ウィンドウタイトルが指定されていません";
                        return false;
                    }
                    hWnd = WindowHelper.FindWindowByTitle(TargetWindowTitle);
                    if (hWnd == IntPtr.Zero)
                    {
                        LogError($"ウィンドウが見つかりません: {TargetWindowTitle}");
                        LastError = $"ウィンドウが見つかりません: {TargetWindowTitle}";
                        return false;
                    }
                    LogInfo($"ウィンドウ発見: {TargetWindowTitle}");
                    break;

                case WindowReferenceMethod.LaunchActionIndex:
                    if (Context == null)
                    {
                        LogError("実行コンテキストが初期化されていません");
                        LastError = "実行コンテキストが初期化されていません";
                        return false;
                    }
                    hWnd = WindowHelper.FindWindowByLaunchAction(Context, TargetLaunchActionIndex);
                    if (hWnd == IntPtr.Zero)
                    {
                        LogError($"起動アクション #{TargetLaunchActionIndex} のウィンドウが見つかりません");
                        LastError = $"起動アクション #{TargetLaunchActionIndex} のウィンドウが見つかりません";
                        return false;
                    }
                    LogInfo($"起動アクション #{TargetLaunchActionIndex} のウィンドウを使用");
                    break;
            }

            // ウィンドウの矩形を取得
            if (NativeMethods.GetWindowRect(hWnd, out var rect))
            {
                searchArea = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
                LogInfo($"ウィンドウ範囲で検索: X={rect.Left}, Y={rect.Top}, W={rect.Right - rect.Left}, H={rect.Bottom - rect.Top}");
            }
            else
            {
                LogWarn("ウィンドウ矩形の取得に失敗しました。全画面で検索します。");
            }
        }
        // 手動指定の検索範囲がある場合
        else if (SearchAreaWidth > 0 && SearchAreaHeight > 0)
        {
            searchArea = new Rectangle(SearchAreaX, SearchAreaY, SearchAreaWidth, SearchAreaHeight);
            LogInfo($"手動指定範囲で検索: X={SearchAreaX}, Y={SearchAreaY}, W={SearchAreaWidth}, H={SearchAreaHeight}");
        }
        // どちらもない場合は全画面検索
        else
        {
            LogInfo("全画面で検索");
        }

        using var matchingService = new ImageMatchingService();

        var location = await matchingService.FindImageWithTimeoutAsync(
            TemplateImagePath!,
            MatchThreshold,
            SearchTimeoutMs,
            UseMultiScale,
            searchArea
        );

        if (location.HasValue)
        {
            LogInfo($"画像発見: ({location.Value.X}, {location.Value.Y})");

            // カーソル移動
            NativeMethods.SetCursorPos(location.Value.X, location.Value.Y);
            await Task.Delay(50);

            // クリック実行
            await PerformClick(ClickType);

            await Task.Delay(DelayAfterClick);
            return true;
        }

        LogError($"画像が見つかりませんでした: {Path.GetFileName(TemplateImagePath)}");
        LastError = $"タイムアウト: {SearchTimeoutMs}ms以内に画像が見つかりませんでした";
        return false;
    }

    // === クリック実行 ===

    private async Task PerformClick(MouseClickType type)
    {
        switch (type)
        {
            case MouseClickType.LeftClick:
                PerformLeftClick();
                break;
            case MouseClickType.RightClick:
                PerformRightClick();
                break;
            case MouseClickType.DoubleClick:
                PerformLeftClick();
                await Task.Delay(50);
                PerformLeftClick();
                break;
            case MouseClickType.MiddleClick:
                PerformMiddleClick();
                break;
        }
    }

    private void PerformLeftClick()
    {
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
    }

    private void PerformRightClick()
    {
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
    }

    private void PerformMiddleClick()
    {
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, UIntPtr.Zero);
        NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_MIDDLEUP, 0, 0, 0, UIntPtr.Zero);
    }

    private string GetClickTypeName(MouseClickType type)
    {
        return type switch
        {
            MouseClickType.LeftClick => "左クリック",
            MouseClickType.RightClick => "右クリック",
            MouseClickType.DoubleClick => "ダブルクリック",
            MouseClickType.MiddleClick => "中央クリック",
            _ => "クリック"
        };
    }
}
