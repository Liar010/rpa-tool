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

    public override string Description
    {
        get
        {
            var desc = $"ファイル: {(string.IsNullOrEmpty(FilePath) ? $"#{OpenActionIndex}" : FilePath)}, " +
                       $"シート: {(string.IsNullOrEmpty(SheetName) ? "(最初のシート)" : SheetName)}, ";

            if (ReadVariableRows)
            {
                desc += $"可変行テーブル（開始列: {StartColumn}, ヘッダー行: {HeaderRow}, データ開始行: {DataStartRow}）";
            }
            else
            {
                desc += $"範囲: {Range}";
            }

            return desc;
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
    /// 範囲（例: A1:B10）※固定範囲モードの場合のみ使用
    /// </summary>
    public string Range { get; set; } = string.Empty;

    /// <summary>
    /// 可変行数読み取りモード（テーブルのヘッダー固定、データ行数可変）
    /// </summary>
    public bool ReadVariableRows { get; set; } = false;

    /// <summary>
    /// テーブルの開始列（例: "A"）※可変行数モードの場合のみ使用
    /// </summary>
    public string StartColumn { get; set; } = "A";

    /// <summary>
    /// ヘッダー行番号（例: 1）※可変行数モードの場合のみ使用
    /// </summary>
    public int HeaderRow { get; set; } = 1;

    /// <summary>
    /// データ開始行番号（例: 2）※可変行数モードの場合のみ使用
    /// </summary>
    public int DataStartRow { get; set; } = 2;

    /// <summary>
    /// 列数（空の場合は最初の空列まで自動検出）※可変行数モードの場合のみ使用
    /// </summary>
    public int ColumnCount { get; set; } = 0;

    /// <summary>
    /// 読み取り結果（タブ区切り・改行区切り）
    /// </summary>
    [JsonIgnore]
    public string Result { get; set; } = string.Empty;

    /// <summary>
    /// 読み取った行数（可変行数モード用）
    /// </summary>
    [JsonIgnore]
    public int RowCount { get; set; } = 0;

    /// <summary>
    /// 読み取った列数（可変行数モード用）
    /// </summary>
    [JsonIgnore]
    public int ActualColumnCount { get; set; } = 0;

    public override bool Validate()
    {
        if (string.IsNullOrEmpty(FilePath) && OpenActionIndex <= 0)
        {
            LogError("ファイルパスまたは開くアクション番号を指定してください");
            return false;
        }

        if (ReadVariableRows)
        {
            // 可変行数モードのバリデーション
            if (string.IsNullOrEmpty(StartColumn))
            {
                LogError("開始列を指定してください");
                return false;
            }

            if (HeaderRow <= 0)
            {
                LogError("ヘッダー行番号は1以上である必要があります");
                return false;
            }

            if (DataStartRow <= HeaderRow)
            {
                LogError("データ開始行はヘッダー行より後である必要があります");
                return false;
            }
        }
        else
        {
            // 固定範囲モードのバリデーション
            if (string.IsNullOrEmpty(Range))
            {
                LogError("範囲を指定してください");
                return false;
            }
        }

        return true;
    }

    public override async Task<bool> ExecuteAsync()
    {
        try
        {
            if (ReadVariableRows)
            {
                LogInfo($"可変行テーブル読み取り開始（ヘッダー行: {HeaderRow}, データ開始: {DataStartRow}）");
            }
            else
            {
                LogInfo($"範囲読み取り開始: {Range}");
            }

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

            // モードに応じて読み取り
            if (ReadVariableRows)
            {
                return await ReadVariableRowsData(worksheet);
            }
            else
            {
                return await ReadFixedRange(worksheet);
            }
        }
        catch (Exception ex)
        {
            LogError($"範囲読み取りエラー: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 固定範囲モードでの読み取り
    /// </summary>
    private async Task<bool> ReadFixedRange(IXLWorksheet worksheet)
    {
        try
        {
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
            RowCount = rowCount;
            ActualColumnCount = colCount;

            LogInfo($"範囲読み取り完了: {rowCount}行 x {colCount}列");
            LogDebug($"読み取り結果:\n{Result}");

            return await Task.FromResult(true);
        }
        catch (Exception ex)
        {
            LogError($"固定範囲読み取りエラー: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 可変行数モードでの読み取り
    /// </summary>
    private async Task<bool> ReadVariableRowsData(IXLWorksheet worksheet)
    {
        try
        {
            var sb = new StringBuilder();

            // 列数を決定（指定されていない場合は自動検出）
            int actualColCount = ColumnCount;
            if (actualColCount <= 0)
            {
                // ヘッダー行から列数を自動検出
                actualColCount = DetectColumnCount(worksheet);
                LogDebug($"列数を自動検出: {actualColCount}列");
            }

            // ヘッダー行を読み取り
            var headerRow = ReadRow(worksheet, HeaderRow, StartColumn, actualColCount);
            sb.Append(headerRow);

            int dataRowCount = 0;

            // データ行を読み取り（空行まで）
            for (int row = DataStartRow; ; row++)
            {
                var rowData = ReadRow(worksheet, row, StartColumn, actualColCount);

                // 空行チェック（すべてのセルが空の場合は終了）
                if (IsEmptyRow(rowData))
                {
                    LogDebug($"行 {row} が空行のため読み取りを終了");
                    break;
                }

                sb.Append('\n');
                sb.Append(rowData);
                dataRowCount++;

                // 安全装置: 10000行以上は読み取らない
                if (dataRowCount >= 10000)
                {
                    LogWarn("10000行に達したため読み取りを打ち切りました");
                    break;
                }
            }

            Result = sb.ToString();
            RowCount = dataRowCount + 1; // ヘッダー行を含む
            ActualColumnCount = actualColCount;

            LogInfo($"可変行テーブル読み取り完了: {RowCount}行（ヘッダー1 + データ{dataRowCount}）x {actualColCount}列");
            LogDebug($"読み取り結果:\n{Result}");

            return await Task.FromResult(true);
        }
        catch (Exception ex)
        {
            LogError($"可変行読み取りエラー: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 列数を自動検出（ヘッダー行から最初の空セルまで）
    /// </summary>
    private int DetectColumnCount(IXLWorksheet worksheet)
    {
        int colIndex = ColumnLetterToNumber(StartColumn);
        int count = 0;

        for (int i = 0; i < 100; i++) // 最大100列まで
        {
            var cell = worksheet.Cell(HeaderRow, colIndex + i);
            if (cell.IsEmpty())
                break;

            count++;
        }

        return count > 0 ? count : 1; // 最低1列
    }

    /// <summary>
    /// 指定行を読み取り（タブ区切り文字列で返す）
    /// </summary>
    private string ReadRow(IXLWorksheet worksheet, int rowNum, string startCol, int colCount)
    {
        var sb = new StringBuilder();
        int startColIndex = ColumnLetterToNumber(startCol);

        for (int i = 0; i < colCount; i++)
        {
            if (i > 0)
                sb.Append('\t');

            var cell = worksheet.Cell(rowNum, startColIndex + i);
            string cellText = cell.Value.ToString() ?? string.Empty;
            sb.Append(cellText);
        }

        return sb.ToString();
    }

    /// <summary>
    /// 行データが空かどうかをチェック
    /// </summary>
    private bool IsEmptyRow(string rowData)
    {
        if (string.IsNullOrWhiteSpace(rowData))
            return true;

        // タブで分割してすべて空白かチェック
        var cells = rowData.Split('\t');
        return cells.All(cell => string.IsNullOrWhiteSpace(cell));
    }

    /// <summary>
    /// 列文字（A, B, ... Z, AA, AB, ...）を数値インデックス（1始まり）に変換
    /// </summary>
    private int ColumnLetterToNumber(string columnLetter)
    {
        columnLetter = columnLetter.ToUpper();
        int result = 0;

        for (int i = 0; i < columnLetter.Length; i++)
        {
            result *= 26;
            result += (columnLetter[i] - 'A' + 1);
        }

        return result;
    }
}
