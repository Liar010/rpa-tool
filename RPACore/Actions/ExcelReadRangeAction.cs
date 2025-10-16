using System.Text;
using System.Text.Json.Serialization;
using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelの範囲を読み取るアクション
/// </summary>
public class ExcelReadRangeAction : ActionBase
{
    public override string Name => "Excel範囲読み取り";

    public override string Description =>
        $"ファイル: {(string.IsNullOrEmpty(FilePath) ? $"#{OpenActionIndex}" : FilePath)}, " +
        $"シート: {(string.IsNullOrEmpty(SheetName) ? "(最初のシート)" : SheetName)}, " +
        $"範囲: {Range}";

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
    /// 読み取り結果（タブ区切り・改行区切り）
    /// </summary>
    [JsonIgnore]
    public string Result { get; set; } = string.Empty;

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

        return true;
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            LogInfo($"範囲読み取り開始: {Range}");

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

            // 範囲を読み取り
            var range = worksheet.Range(Range);
            var sb = new StringBuilder();

            int rowCount = 0;
            int colCount = 0;

            foreach (var row in range.Rows())
            {
                if (rowCount > 0)
                    sb.Append('\n');

                int currentCol = 0;
                foreach (var cell in row.Cells())
                {
                    if (currentCol > 0)
                        sb.Append('\t');

                    string cellText = cell.Value.ToString() ?? string.Empty;
                    sb.Append(cellText);

                    currentCol++;
                }

                if (rowCount == 0)
                    colCount = currentCol;

                rowCount++;
            }

            Result = sb.ToString();

            LogInfo($"範囲読み取り完了: {rowCount}行 x {colCount}列");
            LogDebug($"読み取り結果:\n{Result}");

            return await Task.FromResult(true);
        }
        catch (Exception ex)
        {
            LogError($"範囲読み取りエラー: {ex.Message}");
            return false;
        }
    }
}
