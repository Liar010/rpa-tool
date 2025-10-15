using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelシート名を変更するアクション
/// </summary>
public class ExcelRenameSheetAction : ActionBase
{
    /// <summary>対象のExcelファイルパス（直接指定の場合）</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>開くアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int OpenActionIndex { get; set; } = 0;

    /// <summary>変更前のシート名</summary>
    public string OldSheetName { get; set; } = string.Empty;

    /// <summary>変更後のシート名</summary>
    public string NewSheetName { get; set; } = string.Empty;

    public override string Name => "Excel: シート名を変更";

    public override string Description
    {
        get
        {
            string fileRef = OpenActionIndex > 0 ? $"#{OpenActionIndex}で開いたファイル" : FilePath;
            return $"シート '{OldSheetName}' を '{NewSheetName}' に変更: {fileRef}";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(OldSheetName))
            {
                LogError("変更前のシート名が指定されていません");
                return ContinueOnError;
            }

            if (string.IsNullOrWhiteSpace(NewSheetName))
            {
                LogError("変更後のシート名が指定されていません");
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
                // 変更前のシートの存在確認
                if (!workbook.Worksheets.TryGetWorksheet(OldSheetName, out var worksheet))
                {
                    throw new InvalidOperationException($"シート '{OldSheetName}' が見つかりません");
                }

                // 新しいシート名の重複チェック
                if (workbook.Worksheets.Any(ws => ws.Name == NewSheetName))
                {
                    throw new InvalidOperationException($"シート '{NewSheetName}' は既に存在します");
                }

                // シート名を変更
                worksheet.Name = NewSheetName;

                LogInfo($"シート名を '{OldSheetName}' から '{NewSheetName}' に変更しました");
            });

            return true;
        }
        catch (Exception ex)
        {
            LogError($"シート名の変更中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
