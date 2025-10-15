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
/// </summary>
public class MouseAction : ActionBase
{
    public int X { get; set; }
    public int Y { get; set; }
    public MouseClickType ClickType { get; set; } = MouseClickType.LeftClick;
    public int DelayAfterClick { get; set; } = 100; // クリック後の待機時間（ミリ秒）

    public MouseAction()
    {
        Name = "マウスクリック";
        Description = "指定座標でマウスをクリックします";
    }

    public MouseAction(int x, int y, MouseClickType clickType = MouseClickType.LeftClick)
    {
        Name = "マウスクリック";
        Description = $"座標 ({x}, {y}) で{GetClickTypeName(clickType)}";
        X = x;
        Y = y;
        ClickType = clickType;
    }

    public override bool Validate()
    {
        if (X < 0 || Y < 0)
        {
            LogError("座標が負の値です");
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

            LogInfo($"座標 ({X}, {Y}) に移動して{GetClickTypeName(ClickType)}");

            // カーソル移動
            NativeMethods.SetCursorPos(X, Y);
            await Task.Delay(50); // カーソル移動後の安定待ち

            // クリック実行
            switch (ClickType)
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

            await Task.Delay(DelayAfterClick);
            return true;
        }
        catch (Exception ex)
        {
            LogError($"マウス操作エラー: {ex.Message}");
            return false;
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
