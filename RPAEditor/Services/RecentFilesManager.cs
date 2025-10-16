using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace RPAEditor.Services;

/// <summary>
/// 最近使ったファイルの履歴を管理するクラス
/// </summary>
public class RecentFilesManager
{
    private const int MaxRecentFiles = 10;
    private readonly string _settingsFilePath;
    private List<RecentFileEntry> _recentFiles = new();

    public RecentFilesManager()
    {
        var appDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "RPAEditor"
        );

        if (!Directory.Exists(appDataPath))
        {
            Directory.CreateDirectory(appDataPath);
        }

        _settingsFilePath = Path.Combine(appDataPath, "recent-files.json");
        LoadRecentFiles();
    }

    /// <summary>
    /// ファイルを履歴に追加（既存の場合は最新に移動）
    /// </summary>
    public void AddFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return;

        try
        {
            // 絶対パスに正規化
            var absolutePath = Path.GetFullPath(filePath);

            // 既存のエントリを削除
            _recentFiles.RemoveAll(f =>
                string.Equals(f.FilePath, absolutePath, StringComparison.OrdinalIgnoreCase));

            // 新しいエントリを先頭に追加
            _recentFiles.Insert(0, new RecentFileEntry
            {
                FilePath = absolutePath,
                LastAccessTime = DateTime.Now
            });

            // 最大件数を超えたら古いものを削除
            if (_recentFiles.Count > MaxRecentFiles)
            {
                _recentFiles = _recentFiles.Take(MaxRecentFiles).ToList();
            }

            SaveRecentFiles();
        }
        catch (Exception ex)
        {
            // パスの正規化やファイル保存に失敗しても続行
            Console.WriteLine($"Failed to add recent file: {ex.Message}");
        }
    }

    /// <summary>
    /// 最近使ったファイルのリストを取得（存在するファイルのみ）
    /// </summary>
    public List<RecentFileInfo> GetRecentFiles()
    {
        var validFiles = new List<RecentFileInfo>();

        // 存在しないファイルを除外しながらリストを作成
        var filesToRemove = new List<RecentFileEntry>();

        foreach (var entry in _recentFiles)
        {
            try
            {
                if (File.Exists(entry.FilePath))
                {
                    var fileInfo = new FileInfo(entry.FilePath);
                    validFiles.Add(new RecentFileInfo
                    {
                        FilePath = entry.FilePath,
                        FileName = Path.GetFileNameWithoutExtension(fileInfo.Name),
                        LastAccessTime = entry.LastAccessTime,
                        FileSize = fileInfo.Length
                    });
                }
                else
                {
                    // 存在しないファイルは削除候補に追加
                    filesToRemove.Add(entry);
                }
            }
            catch (Exception ex)
            {
                // ファイル情報の取得に失敗した場合も削除候補に追加
                Console.WriteLine($"Failed to access file {entry.FilePath}: {ex.Message}");
                filesToRemove.Add(entry);
            }
        }

        // 存在しないファイルを履歴から削除
        if (filesToRemove.Any())
        {
            foreach (var entry in filesToRemove)
            {
                _recentFiles.Remove(entry);
            }
            SaveRecentFiles();
        }

        return validFiles;
    }

    /// <summary>
    /// 履歴をクリア
    /// </summary>
    public void ClearHistory()
    {
        _recentFiles.Clear();
        SaveRecentFiles();
    }

    private void LoadRecentFiles()
    {
        try
        {
            if (File.Exists(_settingsFilePath))
            {
                var json = File.ReadAllText(_settingsFilePath);
                var loaded = JsonSerializer.Deserialize<List<RecentFileEntry>>(json);
                if (loaded != null)
                {
                    _recentFiles = loaded;
                }
            }
        }
        catch (Exception ex)
        {
            // 読み込み失敗時は空のリストで開始
            Console.WriteLine($"Failed to load recent files: {ex.Message}");
            _recentFiles = new List<RecentFileEntry>();
        }
    }

    private void SaveRecentFiles()
    {
        try
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var json = JsonSerializer.Serialize(_recentFiles, options);
            File.WriteAllText(_settingsFilePath, json, System.Text.Encoding.UTF8);
        }
        catch (Exception ex)
        {
            // 保存失敗してもアプリケーションは続行
            Console.WriteLine($"Failed to save recent files: {ex.Message}");
        }
    }
}

/// <summary>
/// 最近使ったファイルのエントリ（保存用）
/// </summary>
internal class RecentFileEntry
{
    public string FilePath { get; set; } = string.Empty;
    public DateTime LastAccessTime { get; set; }
}

/// <summary>
/// 最近使ったファイルの情報（表示用）
/// </summary>
public class RecentFileInfo
{
    public string FilePath { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public DateTime LastAccessTime { get; set; }
    public long FileSize { get; set; }

    public string DisplayName => FileName;
    public string DisplayPath => FilePath;
    public string DisplayTime => LastAccessTime.ToString("yyyy/MM/dd HH:mm");
}
