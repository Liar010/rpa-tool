using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelシートを削除するアクション
/// </summary>
public class ExcelDeleteSheetAction : ActionBase
{
    /// <summary>対象のExcelファイルパス（直接指定の場合）</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>開くアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int OpenActionIndex { get; set; } = 0;

    /// <summary>削除するシート名</summary>
    public string SheetName { get; set; } = string.Empty;

    public override string Name => "Excel: シートを削除";

    public override string Description
    {
        get
        {
            string fileRef = OpenActionIndex > 0 ? $"#{OpenActionIndex}で開いたファイル" : FilePath;
            return $"シート '{SheetName}' を削除: {fileRef}";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(SheetName))
            {
                LogError("シート名が指定されていません");
                return ContinueOnError;
            }

            // ファイルパスを解決
            string resolvedFilePath;
            if (OpenActionIndex > 0)
            {
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
                // シートの存在確認
                if (!workbook.Worksheets.TryGetWorksheet(SheetName, out var worksheet))
                {
                    throw new InvalidOperationException($"シート '{SheetName}' が見つかりません");
                }

                // 最後のシートは削除できない
                if (workbook.Worksheets.Count == 1)
                {
                    throw new InvalidOperationException("最後のシートは削除できません");
                }

                // シートを削除
                worksheet.Delete();

                LogInfo($"シート '{SheetName}' を削除しました");
            });

            return true;
        }
        catch (Exception ex)
        {
            LogError($"シートの削除中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
