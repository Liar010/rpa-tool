using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace RPAEditor.Services;

/// <summary>
/// ファイル名テンプレート変数を展開するヘルパークラス
/// </summary>
public static class FileNameTemplateHelper
{
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

        // カスタム変数を置換
        var customVariables = SettingsManager.Instance.CustomVariables;
        foreach (var kvp in customVariables)
        {
            result = result.Replace($"%{kvp.Key}%", kvp.Value);
        }

        // ファイル名として無効な文字を置換
        result = SanitizeFileName(result);

        return result;
    }

    /// <summary>
    /// ファイル名として使用できない文字を安全な文字に置換
    /// </summary>
    private static string SanitizeFileName(string fileName)
    {
        // Windows/Linuxで使用できない文字: \ / : * ? " < > |
        var invalidChars = new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };

        foreach (var c in invalidChars)
        {
            fileName = fileName.Replace(c, '_');
        }

        return fileName;
    }

    /// <summary>
    /// 利用可能な組み込み変数の一覧を取得
    /// </summary>
    public static List<VariableInfo> GetBuiltInVariables()
    {
        var now = DateTime.Now;
        return new List<VariableInfo>
        {
            new() { Name = "%date%", Description = "日付 (yyyy-MM-dd)", Example = now.ToString("yyyy-MM-dd") },
            new() { Name = "%time%", Description = "時刻 (HH-mm-ss)", Example = now.ToString("HH-mm-ss") },
            new() { Name = "%datetime%", Description = "日時 (yyyy-MM-dd_HH-mm-ss)", Example = now.ToString("yyyy-MM-dd_HH-mm-ss") },
            new() { Name = "%timestamp%", Description = "タイムスタンプ (yyyyMMddHHmmss)", Example = now.ToString("yyyyMMddHHmmss") },
            new() { Name = "%year%", Description = "年 (yyyy)", Example = now.ToString("yyyy") },
            new() { Name = "%month%", Description = "月 (MM)", Example = now.ToString("MM") },
            new() { Name = "%day%", Description = "日 (dd)", Example = now.ToString("dd") },
            new() { Name = "%hour%", Description = "時 (HH)", Example = now.ToString("HH") },
            new() { Name = "%minute%", Description = "分 (mm)", Example = now.ToString("mm") },
            new() { Name = "%second%", Description = "秒 (ss)", Example = now.ToString("ss") },
            new() { Name = "%user%", Description = "ユーザー名", Example = Environment.UserName },
            new() { Name = "%computer%", Description = "コンピューター名", Example = Environment.MachineName }
        };
    }

    /// <summary>
    /// プレビューテキストを生成
    /// </summary>
    public static string GetPreview(string template)
    {
        if (string.IsNullOrWhiteSpace(template))
            return "";

        return Expand(template);
    }
}

/// <summary>
/// 変数情報
/// </summary>
public class VariableInfo
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Example { get; set; } = string.Empty;
}
