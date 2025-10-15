using System.Runtime.InteropServices;

namespace RPACore.Actions;

/// <summary>
/// キーボード操作の種類
/// </summary>
public enum KeyboardActionType
{
    TypeText,      // テキスト入力
    PressKey,      // キー押下（単発）
    HotKey         // ショートカットキー（Ctrl+C など）
}

/// <summary>
/// キーボード操作アクション
/// </summary>
public class KeyboardAction : ActionBase
{
    public KeyboardActionType ActionType { get; set; } = KeyboardActionType.TypeText;
    public string Text { get; set; } = string.Empty;
    public VirtualKey Key { get; set; } = VirtualKey.Enter;
    public VirtualKey[] ModifierKeys { get; set; } = Array.Empty<VirtualKey>();
    public int DelayBetweenKeys { get; set; } = 10; // キー入力間隔（ミリ秒）

    public KeyboardAction()
    {
        Name = "キーボード入力";
        Description = "テキストやキーを入力します";
    }

    public KeyboardAction(string text)
    {
        Name = "テキスト入力";
        Description = $"テキスト \"{text}\" を入力";
        ActionType = KeyboardActionType.TypeText;
        Text = text;
    }

    public KeyboardAction(VirtualKey key, params VirtualKey[] modifiers)
    {
        Name = "キー入力";
        Description = GetKeyDescription(key, modifiers);
        ActionType = modifiers.Length > 0 ? KeyboardActionType.HotKey : KeyboardActionType.PressKey;
        Key = key;
        ModifierKeys = modifiers;
    }

    public override bool Validate()
    {
        if (ActionType == KeyboardActionType.TypeText && string.IsNullOrEmpty(Text))
        {
            LogError("入力テキストが空です");
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

            switch (ActionType)
            {
                case KeyboardActionType.TypeText:
                    await TypeTextAsync();
                    break;
                case KeyboardActionType.PressKey:
                    await PressKeyAsync(Key);
                    break;
                case KeyboardActionType.HotKey:
                    await PressHotKeyAsync();
                    break;
            }

            return true;
        }
        catch (Exception ex)
        {
            LogError($"キーボード操作エラー: {ex.Message}");
            return false;
        }
    }

    private async Task TypeTextAsync()
    {
        LogInfo($"テキスト入力: {Text}");

        // クリップボード経由で入力（日本語対応）
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            SetClipboardText(Text);
            await Task.Delay(100);

            // Ctrl+V でペースト
            await PressHotKeyAsync(VirtualKey.V, VirtualKey.Control);
        }
    }

    private async Task PressKeyAsync(VirtualKey key)
    {
        LogInfo($"キー入力: {key}");

        byte vk = (byte)key;
        NativeMethods.keybd_event(vk, 0, NativeMethods.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
        await Task.Delay(DelayBetweenKeys);
        NativeMethods.keybd_event(vk, 0, NativeMethods.KEYEVENTF_KEYUP, UIntPtr.Zero);
        await Task.Delay(50);
    }

    private async Task PressHotKeyAsync()
    {
        await PressHotKeyAsync(Key, ModifierKeys);
    }

    private async Task PressHotKeyAsync(VirtualKey key, params VirtualKey[] modifiers)
    {
        LogInfo($"ショートカットキー: {GetKeyDescription(key, modifiers)}");

        // 修飾キーを押す
        foreach (var modifier in modifiers)
        {
            byte vk = (byte)modifier;
            NativeMethods.keybd_event(vk, 0, NativeMethods.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
            await Task.Delay(10);
        }

        // メインキーを押す
        byte mainVk = (byte)key;
        NativeMethods.keybd_event(mainVk, 0, NativeMethods.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
        await Task.Delay(10);
        NativeMethods.keybd_event(mainVk, 0, NativeMethods.KEYEVENTF_KEYUP, UIntPtr.Zero);

        // 修飾キーを離す（逆順）
        for (int i = modifiers.Length - 1; i >= 0; i--)
        {
            byte vk = (byte)modifiers[i];
            NativeMethods.keybd_event(vk, 0, NativeMethods.KEYEVENTF_KEYUP, UIntPtr.Zero);
            await Task.Delay(10);
        }

        await Task.Delay(50);
    }

    private void SetClipboardText(string text)
    {
        IntPtr hGlobal = Marshal.StringToHGlobalUni(text);
        NativeMethods.OpenClipboard(IntPtr.Zero);
        NativeMethods.EmptyClipboard();
        NativeMethods.SetClipboardData(NativeMethods.CF_UNICODETEXT, hGlobal);
        NativeMethods.CloseClipboard();
    }

    private string GetKeyDescription(VirtualKey key, VirtualKey[] modifiers)
    {
        var parts = new List<string>();

        foreach (var mod in modifiers)
        {
            parts.Add(mod.ToString());
        }
        parts.Add(key.ToString());

        return string.Join("+", parts);
    }
}
