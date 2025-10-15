using System.IO;

namespace RPACore.Actions;

/// <summary>
/// ファイル名変更アクション
/// </summary>
public class FileRenameAction : ActionBase
{
    public string SourcePath { get; set; } = string.Empty;
    public string NewName { get; set; } = string.Empty;

    public FileRenameAction()
    {
        Name = "ファイル名変更";
    }

    public override string Description =>
        $"ファイル名変更: {SourcePath} → {NewName}";

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            LogInfo($"ファイル名を変更します: {SourcePath} → {NewName}");

            // ファイルパス検証
            if (string.IsNullOrWhiteSpace(SourcePath))
            {
                LogError("元のファイルパスが指定されていません");
                return ContinueOnError;
            }

            if (string.IsNullOrWhiteSpace(NewName))
            {
                LogError("新しいファイル名が指定されていません");
                return ContinueOnError;
            }

            // ファイル存在確認
            if (!File.Exists(SourcePath))
            {
                LogError($"ファイルが見つかりません: {SourcePath}");
                return ContinueOnError;
            }

            // 新しいパスを構築
            string? directory = Path.GetDirectoryName(SourcePath);
            if (string.IsNullOrEmpty(directory))
            {
                LogError($"ディレクトリパスを取得できませんでした: {SourcePath}");
                return ContinueOnError;
            }

            string newPath = Path.Combine(directory, NewName);

            // 同名ファイルの存在確認
            if (File.Exists(newPath))
            {
                LogError($"新しいファイル名のファイルが既に存在します: {newPath}");
                return ContinueOnError;
            }

            // ファイル名変更（実際にはMove）
            await Task.Run(() => File.Move(SourcePath, newPath));

            LogInfo($"ファイル名を変更しました: {Path.GetFileName(SourcePath)} → {NewName}");
            return true;
        }
        catch (UnauthorizedAccessException ex)
        {
            LogError($"アクセス権限がありません: {ex.Message}");
            return ContinueOnError;
        }
        catch (IOException ex)
        {
            LogError($"ファイル名変更中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
        catch (Exception ex)
        {
            LogError($"予期しないエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
