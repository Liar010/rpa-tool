using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelセルに値を書き込むアクション
/// </summary>
public class ExcelWriteCellAction : ActionBase
{
    /// <summary>書き込み対象のExcelファイルパス（直接指定の場合）</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>開くアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int OpenActionIndex { get; set; } = 0;

    /// <summary>シート名（省略時は最初のシート）</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>セルアドレス（例: A1, B2）</summary>
    public string CellAddress { get; set; } = string.Empty;

    /// <summary>書き込む値（直接指定の場合）</summary>
    public string Value { get; set; } = string.Empty;

    /// <summary>読み取りアクションの番号（アクション番号指定の場合、0=未使用）</summary>
    public int ReadActionIndex { get; set; } = 0;

    public override string Name => "Excel: セル値を書き込む";

    public override string Description
    {
        get
        {
            string fileRef = OpenActionIndex > 0 ? $"#{OpenActionIndex}で開いたファイル" : FilePath;
            string valueRef = ReadActionIndex > 0 ? $"#{ReadActionIndex}で読み取った値" : $"'{Value}'";
            return string.IsNullOrWhiteSpace(SheetName)
                ? $"セル {CellAddress} に {valueRef} を書き込む: {fileRef}"
                : $"シート '{SheetName}' のセル {CellAddress} に {valueRef} を書き込む: {fileRef}";
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

            // 書き込む値を解決
            string resolvedValue;
            if (ReadActionIndex > 0)
            {
                // 読み取りアクションから値を取得
                var readAction = Context?.ScriptEngine?.Actions.ElementAtOrDefault(ReadActionIndex - 1) as ExcelReadCellAction;
                if (readAction == null)
                {
                    LogError($"アクション #{ReadActionIndex} が ExcelReadCellAction ではありません");
                    return ContinueOnError;
                }
                resolvedValue = readAction.Value;
                LogInfo($"アクション #{ReadActionIndex} で読み取った値を使用: '{resolvedValue}'");
            }
            else
            {
                resolvedValue = Value;
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

                // セルに値を書き込む
                var cell = worksheet.Cell(CellAddress);
                cell.Value = resolvedValue;

                LogInfo($"セル {CellAddress} に値を書き込みました: '{resolvedValue}'");
            });

            return true;
        }
        catch (Exception ex)
        {
            LogError($"セルの書き込み中にエラーが発生しました: {ex.Message}");
            return ContinueOnError;
        }
    }
}
