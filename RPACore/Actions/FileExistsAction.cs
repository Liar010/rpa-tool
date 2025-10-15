namespace RPACore.Actions;

/// <summary>
/// ファイル/フォルダの存在確認を行うアクション
/// </summary>
public class FileExistsAction : ActionBase
{
    public string TargetPath { get; set; } = string.Empty;

    /// <summary>
    /// チェック対象のタイプ（File or Folder）
    /// </summary>
    public FileSystemType TargetType { get; set; } = FileSystemType.File;

    /// <summary>
    /// 存在しない場合にエラーとするか（trueの場合、存在しないとエラー）
    /// </summary>
    public bool FailIfNotExists { get; set; } = true;

    public override string Name => "ファイル/フォルダ存在確認";

    public override string Description
    {
        get
        {
            string typeStr = TargetType == FileSystemType.File ? "ファイル" : "フォルダ";
            string actionStr = FailIfNotExists ? "存在確認" : "存在チェック";
            return $"{typeStr}{actionStr}: {TargetPath}";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            // 入力検証
            if (string.IsNullOrWhiteSpace(TargetPath))
            {
                LogError("対象パスが指定されていません");
                return ContinueOnError;
            }

            LogInfo($"{(TargetType == FileSystemType.File ? "ファイル" : "フォルダ")}存在確認: {TargetPath}");

            bool exists = await Task.Run(() =>
            {
                if (TargetType == FileSystemType.File)
                {
                    return File.Exists(TargetPath);
                }
                else
                {
                    return Directory.Exists(TargetPath);
                }
            });

            if (exists)
            {
                LogInfo($"存在します: {TargetPath}");
                return true;
            }
            else
            {
                if (FailIfNotExists)
                {
                    LogError($"存在しません: {TargetPath}");
                    return ContinueOnError;
                }
                else
                {
                    LogInfo($"存在しません: {TargetPath}");
                    return true; // 存在しなくても成功として扱う
                }
            }
        }
        catch (Exception ex)
        {
            LogError($"存在確認エラー: {ex.Message}");
            return ContinueOnError;
        }
    }
}

/// <summary>
/// ファイルシステムのタイプ
/// </summary>
public enum FileSystemType
{
    File,
    Folder
}
