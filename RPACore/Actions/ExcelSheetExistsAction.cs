using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelシートの存在を確認するアクション
/// </summary>
public class ExcelSheetExistsAction : ActionBase
{
    /// <summary>対象のExcelファイルパス（直接指定の場合）</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>開くアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int OpenActionIndex { get; set; } = 0;

    /// <summary>確認するシート名</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>確認結果（実行後に格納される）</summary>
    public bool Result { get; private set; }

    public override string Name => "Excel: シート存在確認";

    public override string Description
    {
        get
        {
            string fileRef = OpenActionIndex > 0 ? $"#{OpenActionIndex}で開いたファイル" : FilePath;
            return $"シート '{SheetName}' の存在確認: {fileRef}";
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
                Result = workbook.Worksheets.TryGetWorksheet(SheetName, out _);

                LogInfo($"シート '{SheetName}' の存在確認: {(Result ? "存在する" : "存在しない")}");
            });

            return true;
        }
        catch (Exception ex)
        {
            LogError($"シートの存在確認中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
