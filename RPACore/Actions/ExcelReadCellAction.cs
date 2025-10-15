using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelセルの値を読み取るアクション
/// </summary>
public class ExcelReadCellAction : ActionBase
{
    /// <summary>読み取り対象のExcelファイルパス（直接指定の場合）</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>開くアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int OpenActionIndex { get; set; } = 0;

    /// <summary>シート名（省略時は最初のシート）</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>セルアドレス（例: A1, B2）</summary>
    public string CellAddress { get; set; } = string.Empty;

    /// <summary>読み取った値（実行後に格納される）</summary>
    public string Value { get; private set; } = string.Empty;

    public override string Name => "Excel: セル値を読み取る";

    public override string Description
    {
        get
        {
            string fileRef = OpenActionIndex > 0 ? $"#{OpenActionIndex}で開いたファイル" : FilePath;
            return string.IsNullOrWhiteSpace(SheetName)
                ? $"セル {CellAddress} を読み取る: {fileRef}"
                : $"シート '{SheetName}' のセル {CellAddress} を読み取る: {fileRef}";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(CellAddress))
            {
                LogError("セルアドレスが指定されていません");
                return ContinueOnError;
            }

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
                // シートを取得
                IXLWorksheet worksheet;
                if (string.IsNullOrWhiteSpace(SheetName))
                {
                    worksheet = workbook.Worksheets.First();
                }
                else
                {
                    if (!workbook.Worksheets.TryGetWorksheet(SheetName, out worksheet!))
                    {
                        throw new InvalidOperationException($"シート '{SheetName}' が見つかりません");
                    }
                }

                // セルの値を読み取る
                var cell = worksheet.Cell(CellAddress);
                Value = cell.GetString();

                LogInfo($"セル {CellAddress} の値を読み取りました: '{Value}'");
            });

            return true;
        }
        catch (Exception ex)
        {
            LogError($"セルの読み取り中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
