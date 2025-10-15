namespace RPACore.Actions;

/// <summary>
/// すべてのRPAアクションの基本インターフェース
/// </summary>
public interface IAction
{
    /// <summary>
    /// アクションの名前（GUI表示用）
    /// </summary>
    string Name { get; }

    /// <summary>
    /// アクションの説明
    /// </summary>
    string Description { get; }

    /// <summary>
    /// アクションを実行する
    /// </summary>
    /// <returns>実行が成功したかどうか</returns>
    Task<bool> ExecuteAsync();

    /// <summary>
    /// アクションの検証（実行前チェック）
    /// </summary>
    /// <returns>検証結果</returns>
    bool Validate();
}
