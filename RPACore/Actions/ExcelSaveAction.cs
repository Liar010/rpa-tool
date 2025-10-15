using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelファイルを保存するアクション
/// </summary>
public class ExcelSaveAction : ActionBase
{
    /// <summary>保存するExcelファイルのパス（直接指定の場合）</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>開くアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int OpenActionIndex { get; set; } = 0;

    /// <summary>別名で保存する場合の新しいパス（省略時は上書き保存）</summary>
    public string SaveAsPath { get; set; } = string.Empty;

    public override string Name => "Excel: ファイルを保存";

    public override string Description
    {
        get
        {
            string fileRef = OpenActionIndex > 0 ? $"#{OpenActionIndex}で開いたファイル" : FilePath;
            return string.IsNullOrWhiteSpace(SaveAsPath)
                ? $"Excelファイルを保存: {fileRef}"
                : $"Excelファイルを別名保存: {fileRef} → {SaveAsPath}";
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
                if (string.IsNullOrWhiteSpace(SaveAsPath))
                {
                    // 上書き保存
                    workbook.Save();
                    LogInfo($"Excelファイルを保存しました: {fullPath}");
                }
                else
                {
                    // 別名保存
                    string newFullPath = Path.GetFullPath(SaveAsPath);
                    workbook.SaveAs(newFullPath);
                    LogInfo($"Excelファイルを別名保存しました: {newFullPath}");

                    // コンテキストを更新（新しいパスでも参照できるようにする）
                    if (Context != null && !Context.OpenWorkbooks.ContainsKey(newFullPath))
                    {
                        Context.OpenWorkbooks[newFullPath] = workbook;
                    }
                }
            });

            return true;
        }
        catch (Exception ex)
        {
            LogError($"Excelファイルの保存中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
