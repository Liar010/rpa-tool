using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelシートをコピーするアクション
/// </summary>
public class ExcelCopySheetAction : ActionBase
{
    /// <summary>対象のExcelファイルパス（直接指定の場合）</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>開くアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int OpenActionIndex { get; set; } = 0;

    /// <summary>コピー元のシート名</summary>
    public string SourceSheetName { get; set; } = string.Empty;

    /// <summary>コピー先のシート名</summary>
    public string DestinationSheetName { get; set; } = string.Empty;

    public override string Name => "Excel: シートをコピー";

    public override string Description
    {
        get
        {
            string fileRef = OpenActionIndex > 0 ? $"#{OpenActionIndex}で開いたファイル" : FilePath;
            return $"シート '{SourceSheetName}' を '{DestinationSheetName}' としてコピー: {fileRef}";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(SourceSheetName))
            {
                LogError("コピー元のシート名が指定されていません");
                return ContinueOnError;
            }

            if (string.IsNullOrWhiteSpace(DestinationSheetName))
            {
                LogError("コピー先のシート名が指定されていません");
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
                // コピー元のシートの存在確認
                if (!workbook.Worksheets.TryGetWorksheet(SourceSheetName, out var sourceWorksheet))
                {
                    throw new InvalidOperationException($"シート '{SourceSheetName}' が見つかりません");
                }

                // コピー先のシート名の重複チェック
                if (workbook.Worksheets.Any(ws => ws.Name == DestinationSheetName))
                {
                    throw new InvalidOperationException($"シート '{DestinationSheetName}' は既に存在します");
                }

                // シートをコピー
                sourceWorksheet.CopyTo(DestinationSheetName);

                LogInfo($"シート '{SourceSheetName}' を '{DestinationSheetName}' としてコピーしました");
            });

            return true;
        }
        catch (Exception ex)
        {
            LogError($"シートのコピー中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
