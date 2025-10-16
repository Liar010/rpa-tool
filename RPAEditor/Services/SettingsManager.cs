using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace RPAEditor.Services;

/// <summary>
/// アプリケーション設定を管理するクラス
/// </summary>
public class SettingsManager
{
    private static SettingsManager? _instance;
    private readonly string _settingsFilePath;
    private AppSettings _settings;

    public static SettingsManager Instance => _instance ??= new SettingsManager();

    private SettingsManager()
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
        _settings = LoadSettings();
    }

    /// <summary>
    /// カスタム変数を取得
    /// </summary>
    public Dictionary<string, string> CustomVariables => _settings.CustomVariables;

    /// <summary>
    /// カスタム変数を追加または更新
    /// </summary>
    public void SetCustomVariable(string key, string value)
    {
        if (string.IsNullOrWhiteSpace(key))
            throw new ArgumentException("変数名は空にできません", nameof(key));

        _settings.CustomVariables[key] = value;
        SaveSettings();
    }

    /// <summary>
    /// カスタム変数を削除
    /// </summary>
    public void RemoveCustomVariable(string key)
    {
        if (_settings.CustomVariables.ContainsKey(key))
        {
            _settings.CustomVariables.Remove(key);
            SaveSettings();
        }
    }

    /// <summary>
    /// すべての設定をリセット
    /// </summary>
    public void ResetToDefault()
    {
        _settings = new AppSettings();
        SaveSettings();
    }

    private AppSettings LoadSettings()
    {
        try
        {
            if (File.Exists(_settingsFilePath))
            {
                var json = File.ReadAllText(_settingsFilePath);
                var loaded = JsonSerializer.Deserialize<AppSettings>(json);
                if (loaded != null)
                {
                    return loaded;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load settings: {ex.Message}");
        }

        return new AppSettings();
    }

    private void SaveSettings()
    {
        try
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var json = JsonSerializer.Serialize(_settings, options);
            File.WriteAllText(_settingsFilePath, json, System.Text.Encoding.UTF8);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save settings: {ex.Message}");
        }
    }
}

/// <summary>
/// アプリケーション設定データ
/// </summary>
public class AppSettings
{
    public Dictionary<string, string> CustomVariables { get; set; } = new()
    {
        // デフォルトのカスタム変数例
        { "project", "MyProject" },
        { "version", "v1.0" }
    };
}
