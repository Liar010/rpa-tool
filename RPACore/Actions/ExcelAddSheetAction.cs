using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelシートを追加するアクション
/// </summary>
public class ExcelAddSheetAction : ActionBase
{
    /// <summary>対象のExcelファイルパス（直接指定の場合）</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>開くアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int OpenActionIndex { get; set; } = 0;

    /// <summary>新しいシート名</summary>
    public string SheetName { get; set; } = string.Empty;

    public override string Name => "Excel: シートを追加";

    public override string Description
    {
        get
        {
            string fileRef = OpenActionIndex > 0 ? $"#{OpenActionIndex}で開いたファイル" : FilePath;
            return $"シート '{SheetName}' を追加: {fileRef}";
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
                // シート名の重複チェック
                if (workbook.Worksheets.Any(ws => ws.Name == SheetName))
                {
                    throw new InvalidOperationException($"シート '{SheetName}' は既に存在します");
                }

                // シートを追加
                workbook.Worksheets.Add(SheetName);

                LogInfo($"シート '{SheetName}' を追加しました");
            });

            return true;
        }
        catch (Exception ex)
        {
            LogError($"シートの追加中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
