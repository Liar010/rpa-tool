using System.Text.Json.Serialization;

namespace RPACore.Actions;

/// <summary>
/// アクションの基底クラス
/// </summary>
public abstract class ActionBase : IAction
{
    [JsonIgnore]
    public virtual string Name { get; protected set; } = string.Empty;

    [JsonIgnore]
    public virtual string Description { get; protected set; } = string.Empty;

    /// <summary>
    /// アクションが有効かどうか
    /// </summary>
    public bool IsEnabled { get; set; } = true;

    /// <summary>
    /// エラー発生時に処理を続行するかどうか
    /// </summary>
    public bool ContinueOnError { get; set; } = false;

    /// <summary>
    /// 最後のエラーメッセージ
    /// </summary>
    [JsonIgnore]
    public string? LastError { get; protected set; }

    /// <summary>
    /// 実行コンテキスト（ScriptEngineから設定される）
    /// </summary>
    [JsonIgnore]
    public ExecutionContext? Context { get; set; }

    /// <summary>
    /// 現在のアクション番号（1始まり、ScriptEngineから設定される）
    /// </summary>
    [JsonIgnore]
    public int ActionIndex { get; set; }

    public abstract Task<bool> ExecuteAsync();

    public virtual bool Validate()
    {
        return true;
    }

    /// <summary>
    /// DEBUGレベルログ出力
    /// </summary>
    protected void LogDebug(string message)
    {
        Logger.Instance.Debug(Name, message);
    }

    /// <summary>
    /// INFOレベルログ出力
    /// </summary>
    protected void LogInfo(string message)
    {
        Logger.Instance.Info(Name, message);
    }

    /// <summary>
    /// WARNレベルログ出力
    /// </summary>
    protected void LogWarn(string message)
    {
        Logger.Instance.Warn(Name, message);
    }

    /// <summary>
    /// ERRORレベルログ出力
    /// </summary>
    protected void LogError(string message)
    {
        LastError = message;
        Logger.Instance.Error(Name, message);
    }
}
