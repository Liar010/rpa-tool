using ClosedXML.Excel;

namespace RPACore.Actions;

/// <summary>
/// Excelファイルを開くアクション
/// </summary>
public class ExcelOpenAction : ActionBase
{
    /// <summary>開くExcelファイルのパス</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>ファイルが存在しない場合に新規作成するか</summary>
    public bool CreateIfNotExists { get; set; } = false;

    public override string Name => "Excel: ファイルを開く";

    public override string Description => CreateIfNotExists
        ? $"Excelファイルを開く: {FilePath} (存在しない場合は作成)"
        : $"Excelファイルを開く: {FilePath}";

    public override Task<bool> ExecuteAsync()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(FilePath))
            {
                LogError("ファイルパスが指定されていません");
                return Task.FromResult(ContinueOnError);
            }

            string fullPath = Path.GetFullPath(FilePath);

            // 既に開いている場合は閉じる
            if (Context?.OpenWorkbooks.ContainsKey(fullPath) == true)
            {
                Context.OpenWorkbooks[fullPath]?.Dispose();
                Context.OpenWorkbooks.Remove(fullPath);
            }

            XLWorkbook workbook;

            if (File.Exists(fullPath))
            {
                // 既存ファイルを開く
                workbook = new XLWorkbook(fullPath);
                LogInfo($"Excelファイルを開きました: {fullPath}");
            }
            else if (CreateIfNotExists)
            {
                // 新規作成
                workbook = new XLWorkbook();
                workbook.AddWorksheet("Sheet1");
                LogInfo($"Excelファイルを新規作成しました: {fullPath}");
            }
            else
            {
                throw new FileNotFoundException($"ファイルが見つかりません: {fullPath}");
            }

            // コンテキストに保存
            Context?.OpenWorkbooks.Add(fullPath, workbook);

            return Task.FromResult(true);
        }
        catch (FileNotFoundException ex)
        {
            LogError($"ファイルが見つかりません: {FilePath}");
            LogError($"詳細: {ex.Message}");
            return Task.FromResult(ContinueOnError);
        }
        catch (IOException ex)
        {
            // ファイルロックエラー（他のプロセスで使用中）
            LogError($"ファイルが他のプロセスで使用中です。Excelで開いている場合は閉じてから実行してください: {FilePath}");
            LogError($"詳細: {ex.Message}");
            return Task.FromResult(ContinueOnError);
        }
        catch (Exception ex)
        {
            LogError($"Excelファイルを開く際にエラーが発生しました: {ex.Message}");
            LogError($"エラー種類: {ex.GetType().Name}");
            return Task.FromResult(ContinueOnError);
        }
    }
}
