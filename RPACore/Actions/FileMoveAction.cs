namespace RPACore.Actions;

/// <summary>
/// ファイル移動アクション
/// </summary>
public class FileMoveAction : ActionBase
{
    /// <summary>移動元ファイルパス</summary>
    public string SourcePath { get; set; } = string.Empty;

    /// <summary>移動先パス（ファイルまたはフォルダ）</summary>
    public string DestinationPath { get; set; } = string.Empty;

    /// <summary>上書きを許可するか</summary>
    public bool Overwrite { get; set; } = false;

    public override string Name => $"ファイル移動: {Path.GetFileName(SourcePath)}";

    public override string Description =>
        $"移動元: {SourcePath}\n" +
        $"移動先: {DestinationPath}\n" +
        $"上書き: {(Overwrite ? "はい" : "いいえ")}";

    public FileMoveAction()
    {
    }

    public FileMoveAction(string sourcePath, string destinationPath, bool overwrite = false)
    {
        SourcePath = sourcePath;
        DestinationPath = destinationPath;
        Overwrite = overwrite;
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            // ソースファイルの存在確認
            if (!File.Exists(SourcePath))
            {
                LogError($"移動元ファイルが見つかりません: {SourcePath}");
                return ContinueOnError;
            }

            // 移動先がフォルダの場合、ファイル名を付加
            string destinationFile = DestinationPath;
            if (Directory.Exists(DestinationPath))
            {
                string fileName = Path.GetFileName(SourcePath);
                destinationFile = Path.Combine(DestinationPath, fileName);
            }

            // 移動先のディレクトリが存在しない場合は作成
            string? destinationDir = Path.GetDirectoryName(destinationFile);
            if (!string.IsNullOrEmpty(destinationDir) && !Directory.Exists(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
                LogInfo($"移動先フォルダを作成しました: {destinationDir}");
            }

            // 上書き確認
            if (File.Exists(destinationFile))
            {
                if (!Overwrite)
                {
                    LogError($"移動先ファイルが既に存在します（上書き無効）: {destinationFile}");
                    return ContinueOnError;
                }
                // 上書きの場合は先に削除
                File.Delete(destinationFile);
            }

            // ファイル移動実行
            File.Move(SourcePath, destinationFile);
            LogInfo($"ファイルを移動しました: {SourcePath} → {destinationFile}");

            await Task.CompletedTask;
            return true;
        }
        catch (Exception ex)
        {
            LogError($"ファイル移動エラー: {ex.Message}");
            return ContinueOnError;
        }
    }

    public override bool Validate()
    {
        if (string.IsNullOrWhiteSpace(SourcePath))
        {
            LogError("移動元ファイルパスが指定されていません");
            return false;
        }

        if (string.IsNullOrWhiteSpace(DestinationPath))
        {
            LogError("移動先パスが指定されていません");
            return false;
        }

        return true;
    }
}
