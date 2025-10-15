namespace RPACore.Actions;

/// <summary>
/// フォルダを作成するアクション
/// </summary>
public class FolderCreateAction : ActionBase
{
    public string FolderPath { get; set; } = string.Empty;

    public override string Name => "フォルダ作成";

    public override string Description => $"フォルダを作成: {FolderPath}";

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            // 入力検証
            if (string.IsNullOrWhiteSpace(FolderPath))
            {
                LogError("フォルダパスが指定されていません");
                return ContinueOnError;
            }

            LogInfo($"フォルダ作成: {FolderPath}");

            // フォルダが既に存在するかチェック
            if (Directory.Exists(FolderPath))
            {
                LogInfo($"フォルダは既に存在します: {FolderPath}");
                return true; // 既に存在する場合は成功として扱う
            }

            // フォルダを作成（親フォルダも自動的に作成される）
            await Task.Run(() => Directory.CreateDirectory(FolderPath));

            LogInfo($"フォルダを作成しました: {FolderPath}");
            return true;
        }
        catch (Exception ex)
        {
            LogError($"フォルダ作成エラー: {ex.Message}");
            return ContinueOnError;
        }
    }
}
