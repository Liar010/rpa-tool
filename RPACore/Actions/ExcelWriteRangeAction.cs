using System.Text;
using System.Text.Json.Serialization;
using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelの範囲に書き込むアクション
/// </summary>
public class ExcelWriteRangeAction : ActionBase
{
    public override string Name => "Excel範囲書き込み";

    public override string Description
    {
        get
        {
            var fileRef = string.IsNullOrEmpty(FilePath) ? $"#{OpenActionIndex}" : FilePath;
            var valueRef = string.IsNullOrEmpty(Value) ? $"#{ReadActionIndex}の値" : "直接入力値";
            return $"ファイル: {fileRef}, シート: {(string.IsNullOrEmpty(SheetName) ? "(最初のシート)" : SheetName)}, 範囲: {Range}, 値: {valueRef}";
        }
    }

    /// <summary>
    /// ファイルパス（直接指定）
    /// </summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>
    /// 開くアクションの番号（#1, #2など）
    /// </summary>
    public int OpenActionIndex { get; set; }

    /// <summary>
    /// シート名（省略時は最初のシート）
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// 範囲（例: A1:B10）
    /// </summary>
    public string Range { get; set; } = string.Empty;

    /// <summary>
    /// 書き込む値（タブ区切り・改行区切り）
    /// </summary>
    public string Value { get; set; } = string.Empty;

    /// <summary>
    /// 読み取りアクションの番号（値を参照する場合）
    /// </summary>
    public int ReadActionIndex { get; set; }

    public override bool Validate()
    {
        if (string.IsNullOrEmpty(FilePath) && OpenActionIndex <= 0)
        {
            LogError("ファイルパスまたは開くアクション番号を指定してください");
            return false;
        }

        if (string.IsNullOrEmpty(Range))
        {
            LogError("範囲を指定してください");
            return false;
        }

        if (string.IsNullOrEmpty(Value) && ReadActionIndex <= 0)
        {
            LogError("書き込む値または読み取りアクション番号を指定してください");
            return false;
        }

        return true;
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            LogInfo($"範囲書き込み開始: {Range}");

            // ワークブックを取得
            XLWorkbook? workbook = null;
            string targetPath = string.Empty;

            if (OpenActionIndex > 0)
            {
                // アクション番号から取得
                if (Context?.ScriptEngine != null)
                {
                    if (OpenActionIndex > Context.ScriptEngine.Actions.Count)
                    {
                        LogError($"アクション #{OpenActionIndex} は存在しません");
                        return false;
                    }

                    var openAction = Context.ScriptEngine.Actions[OpenActionIndex - 1] as ExcelOpenAction;
                    if (openAction == null)
                    {
                        LogError($"アクション #{OpenActionIndex} は ExcelOpenAction ではありません");
                        return false;
                    }
                    targetPath = openAction.FilePath;
                }
            }
            else
            {
                targetPath = FilePath;
            }

            if (string.IsNullOrEmpty(targetPath))
            {
                LogError("ファイルパスが特定できませんでした");
                return false;
            }

            // ExecutionContext から取得
            if (Context?.OpenWorkbooks != null && Context.OpenWorkbooks.ContainsKey(targetPath))
            {
                workbook = Context.OpenWorkbooks[targetPath];
            }
            else
            {
                LogError($"ファイルが開かれていません: {targetPath}");
                return false;
            }

            // シートを取得
            IXLWorksheet worksheet;
            if (string.IsNullOrEmpty(SheetName))
            {
                worksheet = workbook.Worksheets.First();
                LogDebug($"最初のシート '{worksheet.Name}' を使用");
            }
            else
            {
                if (!workbook.Worksheets.Contains(SheetName))
                {
                    LogError($"シート '{SheetName}' が見つかりません");
                    return false;
                }
                worksheet = workbook.Worksheet(SheetName);
            }

            // 書き込む値を取得
            string writeValue = Value;
            if (ReadActionIndex > 0)
            {
                // 読み取りアクションから値を取得
                if (Context?.ScriptEngine != null)
                {
                    if (ReadActionIndex > Context.ScriptEngine.Actions.Count)
                    {
                        LogError($"アクション #{ReadActionIndex} は存在しません");
                        return false;
                    }

                    var readAction = Context.ScriptEngine.Actions[ReadActionIndex - 1] as ExcelReadRangeAction;
                    if (readAction == null)
                    {
                        LogError($"アクション #{ReadActionIndex} は ExcelReadRangeAction ではありません");
                        return false;
                    }
                    writeValue = readAction.Result;
                }
            }

            if (string.IsNullOrEmpty(writeValue))
            {
                LogWarn("書き込む値が空です");
            }

            // 範囲を取得
            var range = worksheet.Range(Range);

            // タブ・改行区切りの文字列を2次元配列に変換して書き込み
            var lines = writeValue.Split('\n');
            int rowIndex = 0;

            foreach (var line in lines)
            {
                var cells = line.Split('\t');
                int colIndex = 0;

                foreach (var cellValue in cells)
                {
                    var currentRow = range.FirstCell().Address.RowNumber + rowIndex;
                    var currentCol = range.FirstCell().Address.ColumnNumber + colIndex;

                    // 範囲外チェック
                    if (currentRow > range.LastCell().Address.RowNumber ||
                        currentCol > range.LastCell().Address.ColumnNumber)
                    {
                        LogWarn($"データが範囲を超えました。行{rowIndex + 1}, 列{colIndex + 1}はスキップされます");
                        colIndex++;
                        continue;
                    }

                    worksheet.Cell(currentRow, currentCol).Value = cellValue;
                    colIndex++;
                }

                rowIndex++;
            }

            LogInfo($"範囲書き込み完了: {rowIndex}行 x {lines.Max(l => l.Split('\t').Length)}列");

            return await Task.FromResult(true);
        }
        catch (Exception ex)
        {
            LogError($"範囲書き込みエラー: {ex.Message}");
            return false;
        }
    }
}
