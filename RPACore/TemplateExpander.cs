using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace RPACore;

/// <summary>
/// テンプレート変数を展開するヘルパークラス
/// </summary>
public static class TemplateExpander
{
    private static Dictionary<string, string>? _customVariables;
    private static readonly string _settingsFilePath;

    static TemplateExpander()
    {
        var appDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "RPAEditor"
        );

        if (!Directory.Exists(appDataPath))
        {
            Directory.CreateDirectory(appDataPath);
        }

        _settingsFilePath = Path.Combine(appDataPath, "settings.json");
    }

    /// <summary>
    /// テンプレート文字列を展開
    /// </summary>
    public static string Expand(string template)
    {
        if (string.IsNullOrWhiteSpace(template))
            return template;

        var result = template;
        var now = DateTime.Now;

        // 組み込み変数を展開
        var builtInVariables = new Dictionary<string, string>
        {
            { "%date%", now.ToString("yyyy-MM-dd") },
            { "%time%", now.ToString("HH-mm-ss") },
            { "%datetime%", now.ToString("yyyy-MM-dd_HH-mm-ss") },
            { "%timestamp%", now.ToString("yyyyMMddHHmmss") },
            { "%user%", Environment.UserName },
            { "%computer%", Environment.MachineName },
            { "%year%", now.ToString("yyyy") },
            { "%month%", now.ToString("MM") },
            { "%day%", now.ToString("dd") },
            { "%hour%", now.ToString("HH") },
            { "%minute%", now.ToString("mm") },
            { "%second%", now.ToString("ss") }
        };

        // 組み込み変数を置換
        foreach (var kvp in builtInVariables)
        {
            result = result.Replace(kvp.Key, kvp.Value);
        }

        // カスタム変数を読み込み（まだ読み込んでいない場合）
        if (_customVariables == null)
        {
            LoadCustomVariables();
        }

        // カスタム変数を置換
        if (_customVariables != null)
        {
            foreach (var kvp in _customVariables)
            {
                result = result.Replace($"%{kvp.Key}%", kvp.Value);
            }
        }

        return result;
    }

    /// <summary>
    /// カスタム変数を設定ファイルから読み込み
    /// </summary>
    private static void LoadCustomVariables()
    {
        try
        {
            if (File.Exists(_settingsFilePath))
            {
                var json = File.ReadAllText(_settingsFilePath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);
                if (settings != null)
                {
                    _customVariables = settings.CustomVariables;
                    return;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load custom variables: {ex.Message}");
        }

        // デフォルト値
        _customVariables = new Dictionary<string, string>();
    }

    /// <summary>
    /// カスタム変数をリロード（設定が変更された場合に呼び出す）
    /// </summary>
    public static void ReloadCustomVariables()
    {
        _customVariables = null;
    }
}

/// <summary>
/// 設定データ構造（RPAEditorのAppSettingsと同じ）
/// </summary>
internal class AppSettings
{
    public Dictionary<string, string> CustomVariables { get; set; } = new();
}
