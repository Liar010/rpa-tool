using System;
using System.Diagnostics;
using System.Threading;

namespace RPACore;

/// <summary>
/// ウィンドウ検索用のヘルパークラス
/// WindowActionとMouseActionで共通利用
/// </summary>
internal static class WindowHelper
{
    /// <summary>
    /// ウィンドウタイトルでウィンドウを検索
    /// </summary>
    public static IntPtr FindWindowByTitle(string title)
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

    /// <summary>
    /// 起動アクションの番号からウィンドウを検索
    /// </summary>
    public static IntPtr FindWindowByLaunchAction(ExecutionContext context, int launchActionIndex)
    {
        if (!context.LaunchedProcesses.TryGetValue(launchActionIndex, out var processInfo))
        {
            return IntPtr.Zero;
        }

        return FindWindowByProcessInfo(processInfo);
    }

    /// <summary>
    /// プロセス情報からウィンドウを検索
    /// </summary>
    private static IntPtr FindWindowByProcessInfo(LaunchedProcessInfo processInfo)
    {
        Process? targetProcess = null;

        try
        {
            targetProcess = Process.GetProcessById(processInfo.ProcessId);
        }
        catch (ArgumentException)
        {
            // プロセスが既に終了している場合
        }

        // アプローチ1: 元のプロセスのMainWindowHandleを試す
        if (targetProcess != null)
        {
            for (int i = 0; i < 10; i++)
            {
                targetProcess.Refresh();
                if (targetProcess.MainWindowHandle != IntPtr.Zero)
                {
                    return targetProcess.MainWindowHandle;
                }
                Thread.Sleep(500);
            }
        }

        // アプローチ2: プロセス名から関連ウィンドウを検索
        var relatedProcesses = Process.GetProcessesByName(processInfo.ProcessName);

        foreach (var relatedProcess in relatedProcesses)
        {
            if (relatedProcess.MainWindowHandle != IntPtr.Zero)
            {
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

            // 同じプロセス名を持つウィンドウを検索
            foreach (var proc in relatedProcesses)
            {
                if (proc.Id == windowProcessId)
                {
                    foundWindow = hWnd;
                    return false; // 見つかったので列挙を停止
                }
            }

            return true; // 次のウィンドウへ
        }, IntPtr.Zero);

        return foundWindow;
    }
}
