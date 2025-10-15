using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RPACore.Actions;

/// <summary>
/// ウィンドウ操作の種類
/// </summary>
public enum WindowActionType
{
    /// <summary>アプリケーション起動</summary>
    Launch,
    /// <summary>ウィンドウをアクティブ化</summary>
    Activate,
    /// <summary>ウィンドウを最大化</summary>
    Maximize,
    /// <summary>ウィンドウを最小化</summary>
    Minimize,
    /// <summary>ウィンドウを通常表示</summary>
    Restore,
    /// <summary>ウィンドウを閉じる</summary>
    Close
}

/// <summary>
/// ウィンドウ参照方法
/// </summary>
public enum WindowReferenceType
{
    /// <summary>ウィンドウタイトルで指定</summary>
    WindowTitle,
    /// <summary>起動アクションの番号で指定</summary>
    LaunchActionIndex
}

/// <summary>
/// ウィンドウ操作アクション
/// </summary>
public class WindowAction : ActionBase
{
    public WindowActionType ActionType { get; set; }

    /// <summary>実行ファイルパス（Launch時に使用）</summary>
    public string ExecutablePath { get; set; } = string.Empty;

    /// <summary>コマンドライン引数（Launch時に使用）</summary>
    public string Arguments { get; set; } = string.Empty;

    /// <summary>ウィンドウ参照方法</summary>
    public WindowReferenceType ReferenceType { get; set; } = WindowReferenceType.WindowTitle;

    /// <summary>ウィンドウタイトル（Activate等で使用、部分一致）</summary>
    public string WindowTitle { get; set; } = string.Empty;

    /// <summary>参照する起動アクションの番号（1始まり）</summary>
    public int LaunchActionIndex { get; set; } = 0;

    /// <summary>起動後の待機時間（ミリ秒）</summary>
    public int WaitAfterLaunch { get; set; } = 1000;

    public override string Name => ActionType switch
    {
        WindowActionType.Launch => $"アプリ起動: {Path.GetFileName(ExecutablePath)}",
        WindowActionType.Activate => GetWindowReferenceName("選択"),
        WindowActionType.Maximize => GetWindowReferenceName("最大化"),
        WindowActionType.Minimize => GetWindowReferenceName("最小化"),
        WindowActionType.Restore => GetWindowReferenceName("元に戻す"),
        WindowActionType.Close => GetWindowReferenceName("閉じる"),
        _ => "ウィンドウ操作"
    };

    public override string Description => ActionType switch
    {
        WindowActionType.Launch => $"実行: {ExecutablePath}\n引数: {Arguments}\n待機: {WaitAfterLaunch}ms",
        WindowActionType.Activate => GetWindowReferenceDescription("アクティブ化"),
        WindowActionType.Maximize => GetWindowReferenceDescription("最大化"),
        WindowActionType.Minimize => GetWindowReferenceDescription("最小化"),
        WindowActionType.Restore => GetWindowReferenceDescription("通常表示"),
        WindowActionType.Close => GetWindowReferenceDescription("閉じる"),
        _ => "ウィンドウ操作"
    };

    private string GetWindowReferenceName(string operation)
    {
        return ReferenceType == WindowReferenceType.LaunchActionIndex
            ? $"{operation}: [アクション #{LaunchActionIndex}]"
            : $"{operation}: {WindowTitle}";
    }

    private string GetWindowReferenceDescription(string operation)
    {
        return ReferenceType == WindowReferenceType.LaunchActionIndex
            ? $"アクション #{LaunchActionIndex} で起動したウィンドウを{operation}"
            : $"タイトル「{WindowTitle}」を含むウィンドウを{operation}";
    }

    public WindowAction()
    {
    }

    public WindowAction(WindowActionType actionType, string windowTitleOrPath)
    {
        ActionType = actionType;
        if (actionType == WindowActionType.Launch)
        {
            ExecutablePath = windowTitleOrPath;
        }
        else
        {
            WindowTitle = windowTitleOrPath;
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            switch (ActionType)
            {
                case WindowActionType.Launch:
                    return await LaunchApplication();

                case WindowActionType.Activate:
                case WindowActionType.Maximize:
                case WindowActionType.Minimize:
                case WindowActionType.Restore:
                case WindowActionType.Close:
                    return await ManipulateWindow();

                default:
                    LogError($"未対応のアクションタイプ: {ActionType}");
                    return false;
            }
        }
        catch (Exception ex)
        {
            LogError($"ウィンドウ操作エラー: {ex.Message}");
            return ContinueOnError;
        }
    }

    private async Task<bool> LaunchApplication()
    {
        if (string.IsNullOrWhiteSpace(ExecutablePath))
        {
            LogError("実行ファイルパスが指定されていません");
            return false;
        }

        if (!File.Exists(ExecutablePath))
        {
            LogError($"ファイルが見つかりません: {ExecutablePath}");
            return false;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = ExecutablePath,
            Arguments = Arguments,
            UseShellExecute = true
        };

        var process = Process.Start(startInfo);

        // プロセス情報をコンテキストに記録
        if (process != null && Context != null)
        {
            Context.LaunchedProcesses[ActionIndex] = new LaunchedProcessInfo
            {
                ProcessId = process.Id,
                ProcessName = process.ProcessName
            };
            LogInfo($"プロセス {process.ProcessName} (PID: {process.Id}) をアクション #{ActionIndex} として記録");
        }

        if (WaitAfterLaunch > 0)
        {
            await Task.Delay(WaitAfterLaunch);
        }

        return true;
    }

    private async Task<bool> ManipulateWindow()
    {
        IntPtr hWnd;

        // 参照方法に応じてウィンドウを検索
        if (ReferenceType == WindowReferenceType.LaunchActionIndex)
        {
            // アクション番号からプロセス情報を取得
            if (Context == null || !Context.LaunchedProcesses.TryGetValue(LaunchActionIndex, out var processInfo))
            {
                LogError($"アクション #{LaunchActionIndex} で起動されたプロセスが見つかりません");
                return false;
            }

            hWnd = FindWindowByProcessInfo(processInfo);

            if (hWnd == IntPtr.Zero)
            {
                LogError($"プロセス {processInfo.ProcessName} (PID: {processInfo.ProcessId}) のウィンドウが見つかりません");
                return false;
            }
        }
        else
        {
            // ウィンドウタイトルで検索
            if (string.IsNullOrWhiteSpace(WindowTitle))
            {
                LogError("ウィンドウタイトルが指定されていません");
                return false;
            }

            hWnd = FindWindowByTitle(WindowTitle);

            if (hWnd == IntPtr.Zero)
            {
                LogError($"ウィンドウが見つかりません: {WindowTitle}");
                return false;
            }
        }

        switch (ActionType)
        {
            case WindowActionType.Activate:
                BringToForeground(hWnd);
                break;

            case WindowActionType.Maximize:
                BringToForeground(hWnd);
                NativeMethods.ShowWindow(hWnd, NativeMethods.SW_MAXIMIZE);
                break;

            case WindowActionType.Minimize:
                NativeMethods.ShowWindow(hWnd, NativeMethods.SW_MINIMIZE);
                break;

            case WindowActionType.Restore:
                BringToForeground(hWnd);
                NativeMethods.ShowWindow(hWnd, NativeMethods.SW_SHOWNORMAL);
                break;

            case WindowActionType.Close:
                NativeMethods.SendMessage(hWnd, NativeMethods.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                break;
        }

        await Task.Delay(100); // ウィンドウ操作の完了を待つ
        return true;
    }

    /// <summary>
    /// ウィンドウを確実に最前面に持ってくる
    /// </summary>
    private void BringToForeground(IntPtr hWnd)
    {
        // 最小化されている場合は復元
        if (NativeMethods.IsIconic(hWnd))
        {
            NativeMethods.ShowWindow(hWnd, NativeMethods.SW_RESTORE);
        }

        // Altキーを押して離すことで、SetForegroundWindowの制限を回避
        NativeMethods.keybd_event(NativeMethods.VK_MENU, 0, NativeMethods.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
        NativeMethods.keybd_event(NativeMethods.VK_MENU, 0, NativeMethods.KEYEVENTF_KEYUP, UIntPtr.Zero);

        // 最前面に持ってくる
        NativeMethods.BringWindowToTop(hWnd);
        NativeMethods.SetForegroundWindow(hWnd);
    }

    private IntPtr FindWindowByTitle(string title)
    {
        IntPtr foundWindow = IntPtr.Zero;

        NativeMethods.EnumWindows((hWnd, lParam) =>
        {
            if (!NativeMethods.IsWindowVisible(hWnd))
                return true;

            int length = NativeMethods.GetWindowTextLength(hWnd);
            if (length == 0)
                return true;

            var sb = new System.Text.StringBuilder(length + 1);
            NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);
            string windowTitle = sb.ToString();

            if (windowTitle.Contains(title, StringComparison.OrdinalIgnoreCase))
            {
                foundWindow = hWnd;
                return false; // 見つかったので列挙を停止
            }

            return true; // 次のウィンドウへ
        }, IntPtr.Zero);

        return foundWindow;
    }

    private IntPtr FindWindowByProcessInfo(LaunchedProcessInfo processInfo)
    {
        Process? targetProcess = null;

        try
        {
            targetProcess = Process.GetProcessById(processInfo.ProcessId);
        }
        catch (ArgumentException)
        {
            // プロセスが既に終了している場合
            LogInfo($"起動プロセス (PID: {processInfo.ProcessId}) は終了しています。プロセス名 '{processInfo.ProcessName}' で検索します");
        }

        // アプローチ1: 元のプロセスのMainWindowHandleを試す（プロセスが生きている場合）
        if (targetProcess != null)
        {
            for (int i = 0; i < 10; i++)
            {
                targetProcess.Refresh();
                if (targetProcess.MainWindowHandle != IntPtr.Zero)
                {
                    LogInfo($"MainWindowHandle で発見: {targetProcess.MainWindowHandle}");
                    return targetProcess.MainWindowHandle;
                }
                Thread.Sleep(500);
            }
            LogInfo("MainWindowHandle が見つからないため、プロセス名で検索します");
        }

        // アプローチ2: プロセス名から関連ウィンドウを検索
        var relatedProcesses = Process.GetProcessesByName(processInfo.ProcessName);

        foreach (var relatedProcess in relatedProcesses)
        {
            if (relatedProcess.MainWindowHandle != IntPtr.Zero)
            {
                LogInfo($"同名プロセス '{processInfo.ProcessName}' のウィンドウを発見: PID {relatedProcess.Id}");
                return relatedProcess.MainWindowHandle;
            }
        }

        // アプローチ3: EnumWindows で直接検索
        IntPtr foundWindow = IntPtr.Zero;
        NativeMethods.EnumWindows((hWnd, lParam) =>
        {
            if (!NativeMethods.IsWindowVisible(hWnd))
                return true;

            NativeMethods.GetWindowThreadProcessId(hWnd, out uint windowProcessId);

            foreach (var rp in relatedProcesses)
            {
                if (windowProcessId == rp.Id)
                {
                    foundWindow = hWnd;
                    return false;
                }
            }

            return true;
        }, IntPtr.Zero);

        if (foundWindow != IntPtr.Zero)
        {
            LogInfo($"EnumWindows で発見: {foundWindow}");
        }
        else
        {
            LogError($"プロセス名 '{processInfo.ProcessName}' のウィンドウが見つかりませんでした");
        }

        return foundWindow;
    }

    public override bool Validate()
    {
        if (ActionType == WindowActionType.Launch)
        {
            return !string.IsNullOrWhiteSpace(ExecutablePath);
        }
        else
        {
            if (ReferenceType == WindowReferenceType.LaunchActionIndex)
            {
                return LaunchActionIndex > 0;
            }
            else
            {
                return !string.IsNullOrWhiteSpace(WindowTitle);
            }
        }
    }
}
