using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelファイルを閉じるアクション
/// </summary>
public class ExcelCloseAction : ActionBase
{
    /// <summary>閉じるExcelファイルのパス（直接指定の場合）</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>開くアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int OpenActionIndex { get; set; } = 0;

    /// <summary>閉じる前に保存するか</summary>
    public bool SaveBeforeClose { get; set; } = true;

    public override string Name => "Excel: ファイルを閉じる";

    public override string Description
    {
        get
        {
            string fileRef = OpenActionIndex > 0 ? $"#{OpenActionIndex}で開いたファイル" : FilePath;
            return SaveBeforeClose
                ? $"Excelファイルを保存して閉じる: {fileRef}"
                : $"Excelファイルを閉じる: {fileRef}";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            // ファイルパスを解決
            string resolvedFilePath;
            if (OpenActionIndex > 0)
            {
                // アクション番号からファイルパスを取得
                var openAction = Context?.ScriptEngine?.Actions.ElementAtOrDefault(OpenActionIndex - 1) as ExcelOpenAction;
                if (openAction == null)
                {
                    LogError($"アクション #{OpenActionIndex} が ExcelOpenAction ではありません");
                    return ContinueOnError;
                }
                resolvedFilePath = openAction.FilePath;
                LogInfo($"アクション #{OpenActionIndex} で開いたファイルを使用: {resolvedFilePath}");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(FilePath))
                {
                    LogError("ファイルパスが指定されていません");
                    return ContinueOnError;
                }
                resolvedFilePath = FilePath;
            }

            string fullPath = Path.GetFullPath(resolvedFilePath);

            // コンテキストからワークブックを取得
            if (Context?.OpenWorkbooks.TryGetValue(fullPath, out var workbook) != true || workbook == null)
            {
                LogError($"Excelファイルが開かれていません: {fullPath}");
                return ContinueOnError;
            }

            await Task.Run(() =>
            {
                if (SaveBeforeClose)
                {
                    workbook.Save();
                    LogInfo($"Excelファイルを保存しました: {fullPath}");
                }

                workbook.Dispose();
                Context?.OpenWorkbooks.Remove(fullPath);

                LogInfo($"Excelファイルを閉じました: {fullPath}");
            });

            return true;
        }
        catch (Exception ex)
        {
            LogError($"Excelファイルを閉じる際にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
