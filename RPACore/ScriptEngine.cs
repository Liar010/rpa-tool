using RPACore.Actions;
using System.Text.Json;
using ClosedXML.Excel;

namespace RPACore;

/// <summary>
/// 起動されたプロセス情報
/// </summary>
public class LaunchedProcessInfo
{
    public int ProcessId { get; set; }
    public string ProcessName { get; set; } = string.Empty;
}

/// <summary>
/// 実行コンテキスト（アクション間でデータを共有）
/// </summary>
public class ExecutionContext
{
    /// <summary>アクション番号（1始まり）→プロセス情報のマッピング</summary>
    public Dictionary<int, LaunchedProcessInfo> LaunchedProcesses { get; } = new();

    /// <summary>開いているExcelワークブック（複数ファイル対応）</summary>
    public Dictionary<string, XLWorkbook> OpenWorkbooks { get; } = new();

    /// <summary>ScriptEngineへの参照（アクションリストにアクセスするため）</summary>
    public ScriptEngine? ScriptEngine { get; set; }

    public void Clear()
    {
        LaunchedProcesses.Clear();

        // 開いているワークブックをすべてクローズ
        foreach (var workbook in OpenWorkbooks.Values)
        {
            workbook?.Dispose();
        }
        OpenWorkbooks.Clear();
    }
}

/// <summary>
/// RPAスクリプトを実行するエンジン
/// </summary>
public class ScriptEngine
{
    public List<IAction> Actions { get; } = new();
    public bool IsRunning { get; private set; }
    public int CurrentActionIndex { get; private set; } = -1;

    /// <summary>実行コンテキスト</summary>
    public ExecutionContext Context { get; } = new();

    public event EventHandler<ActionExecutedEventArgs>? ActionExecuted;
    public event EventHandler<ScriptCompletedEventArgs>? ScriptCompleted;

    /// <summary>
    /// アクションを追加
    /// </summary>
    public void AddAction(IAction action)
    {
        Actions.Add(action);
    }

    /// <summary>
    /// アクションをクリア
    /// </summary>
    public void ClearActions()
    {
        Actions.Clear();
        CurrentActionIndex = -1;
    }

    /// <summary>
    /// スクリプトを実行
    /// </summary>
    public async Task<bool> RunAsync()
    {
        if (IsRunning)
        {
            Console.WriteLine("既にスクリプトが実行中です");
            return false;
        }

        IsRunning = true;
        CurrentActionIndex = 0;
        Context.Clear(); // 実行コンテキストをクリア
        Context.ScriptEngine = this; // ScriptEngineへの参照を設定
        bool allSuccess = true;

        try
        {
            Console.WriteLine($"=== スクリプト実行開始 ({Actions.Count} アクション) ===");

            for (int i = 0; i < Actions.Count; i++)
            {
                CurrentActionIndex = i;
                var action = Actions[i];

                // ActionBaseの場合はContextとActionIndexを設定
                if (action is ActionBase actionBase)
                {
                    actionBase.Context = Context;
                    actionBase.ActionIndex = i + 1; // 1始まり
                }

                Console.WriteLine($"\n[{i + 1}/{Actions.Count}] {action.Name}");

                bool success = await action.ExecuteAsync();

                ActionExecuted?.Invoke(this, new ActionExecutedEventArgs
                {
                    Action = action,
                    Index = i,
                    Success = success
                });

                if (!success)
                {
                    allSuccess = false;

                    // ContinueOnErrorがfalseの場合は中断
                    if (action is ActionBase ab && !ab.ContinueOnError)
                    {
                        Console.WriteLine($"エラーが発生したため、スクリプトを中断します");
                        break;
                    }
                }
            }

            Console.WriteLine($"\n=== スクリプト実行完了 ===");
            ScriptCompleted?.Invoke(this, new ScriptCompletedEventArgs { Success = allSuccess });

            return allSuccess;
        }
        finally
        {
            IsRunning = false;
            CurrentActionIndex = -1;
        }
    }

    /// <summary>
    /// 指定したインデックスまで実行（デバッグ用）
    /// </summary>
    public async Task<bool> RunToIndexAsync(int targetIndex)
    {
        if (targetIndex < 0 || targetIndex >= Actions.Count)
            return false;

        IsRunning = true;
        bool allSuccess = true;

        try
        {
            for (int i = 0; i <= targetIndex; i++)
            {
                CurrentActionIndex = i;
                var action = Actions[i];

                Console.WriteLine($"[{i + 1}/{targetIndex + 1}] {action.Name}");
                bool success = await action.ExecuteAsync();

                ActionExecuted?.Invoke(this, new ActionExecutedEventArgs
                {
                    Action = action,
                    Index = i,
                    Success = success
                });

                if (!success)
                {
                    allSuccess = false;
                    if (action is ActionBase actionBase && !actionBase.ContinueOnError)
                        break;
                }
            }

            return allSuccess;
        }
        finally
        {
            IsRunning = false;
        }
    }

    /// <summary>
    /// スクリプトをJSON形式で保存
    /// </summary>
    public async Task SaveToFileAsync(string filePath)
    {
        var serializerOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            // ランタイムプロパティは無視
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.Never,
            // 日本語をエスケープせずそのまま出力
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        var scriptData = new ScriptData
        {
            Actions = Actions.Select(a => new ActionData
            {
                Type = a.GetType().Name,
                Data = JsonSerializer.Serialize(a, a.GetType(), serializerOptions)
            }).ToList()
        };

        var json = JsonSerializer.Serialize(scriptData, serializerOptions);

        await File.WriteAllTextAsync(filePath, json, System.Text.Encoding.UTF8);
    }

    /// <summary>
    /// スクリプトをJSON形式から読み込み
    /// </summary>
    public async Task LoadFromFileAsync(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"ファイルが見つかりません: {filePath}");
        }

        var json = await File.ReadAllTextAsync(filePath);
        var scriptData = JsonSerializer.Deserialize<ScriptData>(json);

        if (scriptData == null)
        {
            throw new InvalidOperationException("スクリプトファイルの読み込みに失敗しました");
        }

        // 既存のアクションをクリア
        Actions.Clear();

        // 各アクションをデシリアライズして追加
        foreach (var actionData in scriptData.Actions)
        {
            IAction? action = actionData.Type switch
            {
                nameof(MouseAction) => JsonSerializer.Deserialize<MouseAction>(actionData.Data),
                nameof(KeyboardAction) => JsonSerializer.Deserialize<KeyboardAction>(actionData.Data),
                nameof(WaitAction) => JsonSerializer.Deserialize<WaitAction>(actionData.Data),
                nameof(WindowAction) => JsonSerializer.Deserialize<WindowAction>(actionData.Data),
                nameof(FileCopyAction) => JsonSerializer.Deserialize<FileCopyAction>(actionData.Data),
                nameof(FileMoveAction) => JsonSerializer.Deserialize<FileMoveAction>(actionData.Data),
                nameof(FileDeleteAction) => JsonSerializer.Deserialize<FileDeleteAction>(actionData.Data),
                nameof(FileRenameAction) => JsonSerializer.Deserialize<FileRenameAction>(actionData.Data),
                nameof(FolderCreateAction) => JsonSerializer.Deserialize<FolderCreateAction>(actionData.Data),
                nameof(FileExistsAction) => JsonSerializer.Deserialize<FileExistsAction>(actionData.Data),
                nameof(ExcelOpenAction) => JsonSerializer.Deserialize<ExcelOpenAction>(actionData.Data),
                nameof(ExcelReadCellAction) => JsonSerializer.Deserialize<ExcelReadCellAction>(actionData.Data),
                nameof(ExcelWriteCellAction) => JsonSerializer.Deserialize<ExcelWriteCellAction>(actionData.Data),
                nameof(ExcelSaveAction) => JsonSerializer.Deserialize<ExcelSaveAction>(actionData.Data),
                nameof(ExcelCloseAction) => JsonSerializer.Deserialize<ExcelCloseAction>(actionData.Data),
                _ => throw new InvalidOperationException($"未知のアクションタイプ: {actionData.Type}")
            };

            if (action != null)
            {
                Actions.Add(action);
            }
        }

        Console.WriteLine($"スクリプトを読み込みました: {Actions.Count} アクション");
    }
}

public class ActionExecutedEventArgs : EventArgs
{
    public IAction? Action { get; set; }
    public int Index { get; set; }
    public bool Success { get; set; }
}

public class ScriptCompletedEventArgs : EventArgs
{
    public bool Success { get; set; }
}

// 保存用データ構造
internal class ScriptData
{
    public List<ActionData> Actions { get; set; } = new();
}

internal class ActionData
{
    public string Type { get; set; } = string.Empty;
    public string Data { get; set; } = string.Empty;
}
