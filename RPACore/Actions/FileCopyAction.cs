namespace RPACore.Actions;

/// <summary>
/// ファイルコピーアクション
/// </summary>
public class FileCopyAction : ActionBase
{
    /// <summary>コピー元ファイルパス</summary>
    public string SourcePath { get; set; } = string.Empty;

    /// <summary>コピー先パス（ファイルまたはフォルダ）</summary>
    public string DestinationPath { get; set; } = string.Empty;

    /// <summary>上書きを許可するか</summary>
    public bool Overwrite { get; set; } = false;

    public override string Name => $"ファイルコピー: {Path.GetFileName(SourcePath)}";

    public override string Description =>
        $"コピー元: {SourcePath}\n" +
        $"コピー先: {DestinationPath}\n" +
        $"上書き: {(Overwrite ? "はい" : "いいえ")}";

    public FileCopyAction()
    {
    }

    public FileCopyAction(string sourcePath, string destinationPath, bool overwrite = false)
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
                LogError($"コピー元ファイルが見つかりません: {SourcePath}");
                return ContinueOnError;
            }

            // コピー先がフォルダの場合、ファイル名を付加
            string destinationFile = DestinationPath;
            if (Directory.Exists(DestinationPath))
            {
                string fileName = Path.GetFileName(SourcePath);
                destinationFile = Path.Combine(DestinationPath, fileName);
            }

            // コピー先のディレクトリが存在しない場合は作成
            string? destinationDir = Path.GetDirectoryName(destinationFile);
            if (!string.IsNullOrEmpty(destinationDir) && !Directory.Exists(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
                LogInfo($"コピー先フォルダを作成しました: {destinationDir}");
            }

            // 上書き確認
            if (File.Exists(destinationFile) && !Overwrite)
            {
                LogError($"コピー先ファイルが既に存在します（上書き無効）: {destinationFile}");
                return ContinueOnError;
            }

            // ファイルコピー実行
            File.Copy(SourcePath, destinationFile, Overwrite);
            LogInfo($"ファイルをコピーしました: {SourcePath} → {destinationFile}");

            await Task.CompletedTask;
            return true;
        }
        catch (Exception ex)
        {
            LogError($"ファイルコピーエラー: {ex.Message}");
            return ContinueOnError;
        }
    }

    public override bool Validate()
    {
        if (string.IsNullOrWhiteSpace(SourcePath))
        {
            LogError("コピー元ファイルパスが指定されていません");
            return false;
        }

        if (string.IsNullOrWhiteSpace(DestinationPath))
        {
            LogError("コピー先パスが指定されていません");
            return false;
        }

        return true;
    }
}
