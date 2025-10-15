using System.IO;

namespace RPACore.Actions;

/// <summary>
/// ファイル削除アクション
/// </summary>
public class FileDeleteAction : ActionBase
{
    public string FilePath { get; set; } = string.Empty;

    public FileDeleteAction()
    {
        Name = "ファイル削除";
    }

    public override string Description =>
        $"ファイル削除: {FilePath}";

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            LogInfo($"ファイルを削除します: {FilePath}");

            // ファイルパス検証
            if (string.IsNullOrWhiteSpace(FilePath))
            {
                LogError("ファイルパスが指定されていません");
                return ContinueOnError;
            }

            // ファイル存在確認
            if (!File.Exists(FilePath))
            {
                LogError($"ファイルが見つかりません: {FilePath}");
                return ContinueOnError;
            }

            // ファイル削除
            await Task.Run(() => File.Delete(FilePath));

            LogInfo($"ファイルを削除しました: {FilePath}");
            return true;
        }
        catch (UnauthorizedAccessException ex)
        {
            LogError($"アクセス権限がありません: {ex.Message}");
            return ContinueOnError;
        }
        catch (IOException ex)
        {
            LogError($"ファイル削除中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
        catch (Exception ex)
        {
            LogError($"予期しないエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
