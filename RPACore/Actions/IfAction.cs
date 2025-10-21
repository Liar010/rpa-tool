using System;
using System.Threading.Tasks;

namespace RPACore.Actions;

/// <summary>
/// 条件比較の種類
/// </summary>
public enum ConditionType
{
    Equal,              // 等しい (==)
    NotEqual,           // 等しくない (!=)
    GreaterThan,        // より大きい (>)
    GreaterThanOrEqual, // 以上 (>=)
    LessThan,           // より小さい (<)
    LessThanOrEqual,    // 以下 (<=)
    Contains,           // 含む
    NotContains,        // 含まない
    IsEmpty,            // 空である
    IsNotEmpty          // 空でない
}

/// <summary>
/// 条件が真の場合のアクション
/// </summary>
public enum IfThenAction
{
    Continue,           // 次のアクションに進む
    SkipNext,           // 次のアクションをスキップ
    JumpToAction,       // 指定したアクション番号にジャンプ
    ExitScript          // スクリプトを終了
}

/// <summary>
/// 条件分岐アクション
/// </summary>
public class IfAction : ActionBase
{
    /// <summary>比較する左辺値（変数名や固定値）</summary>
    public string LeftValue { get; set; } = string.Empty;

    /// <summary>比較する右辺値（変数名や固定値）</summary>
    public string RightValue { get; set; } = string.Empty;

    /// <summary>条件の種類</summary>
    public ConditionType ConditionType { get; set; } = ConditionType.Equal;

    /// <summary>条件が真の場合のアクション</summary>
    public IfThenAction ThenAction { get; set; } = IfThenAction.Continue;

    /// <summary>ジャンプ先のアクション番号（ThenAction=JumpToActionの場合）</summary>
    public int JumpToActionIndex { get; set; } = 0;

    public override string Name => "条件分岐 (If)";

    public override string Description
    {
        get
        {
            var conditionStr = ConditionType switch
            {
                ConditionType.Equal => $"{LeftValue} == {RightValue}",
                ConditionType.NotEqual => $"{LeftValue} != {RightValue}",
                ConditionType.GreaterThan => $"{LeftValue} > {RightValue}",
                ConditionType.GreaterThanOrEqual => $"{LeftValue} >= {RightValue}",
                ConditionType.LessThan => $"{LeftValue} < {RightValue}",
                ConditionType.LessThanOrEqual => $"{LeftValue} <= {RightValue}",
                ConditionType.Contains => $"{LeftValue} に '{RightValue}' を含む",
                ConditionType.NotContains => $"{LeftValue} に '{RightValue}' を含まない",
                ConditionType.IsEmpty => $"{LeftValue} が空",
                ConditionType.IsNotEmpty => $"{LeftValue} が空でない",
                _ => $"{LeftValue} ? {RightValue}"
            };

            var actionStr = ThenAction switch
            {
                IfThenAction.Continue => "次へ",
                IfThenAction.SkipNext => "次をスキップ",
                IfThenAction.JumpToAction => $"#{JumpToActionIndex}へジャンプ",
                IfThenAction.ExitScript => "終了",
                _ => "?"
            };

            return $"If ({conditionStr}) Then {actionStr}";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        await Task.CompletedTask;

        try
        {
            if (Context == null)
            {
                LogError("実行コンテキストが利用できません");
                return false;
            }

            // 条件を評価
            bool conditionResult = EvaluateCondition();

            LogInfo($"条件評価: {Description} → {(conditionResult ? "真" : "偽")}");

            if (conditionResult)
            {
                // 条件が真の場合のアクションを実行
                switch (ThenAction)
                {
                    case IfThenAction.Continue:
                        // 何もしない（次のアクションに進む）
                        break;

                    case IfThenAction.SkipNext:
                        // ScriptEngineに次のアクションをスキップするよう指示
                        // これは ScriptEngine 側で実装が必要
                        LogInfo("次のアクションをスキップします");
                        // TODO: ScriptEngine でスキップ機能を実装
                        break;

                    case IfThenAction.JumpToAction:
                        if (JumpToActionIndex > 0 && JumpToActionIndex <= (Context.ScriptEngine?.Actions.Count ?? 0))
                        {
                            LogInfo($"アクション #{JumpToActionIndex} にジャンプします");
                            // TODO: ScriptEngine でジャンプ機能を実装
                        }
                        else
                        {
                            LogError($"無効なジャンプ先: #{JumpToActionIndex}");
                            return false;
                        }
                        break;

                    case IfThenAction.ExitScript:
                        LogInfo("スクリプトを終了します");
                        // TODO: ScriptEngine で終了機能を実装
                        break;
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            LogError($"条件分岐中にエラーが発生しました: {ex.Message}");
            return false;
        }
    }

    private bool EvaluateCondition()
    {
        // テンプレート変数展開は ScriptEngine で自動的に行われているため、
        // ここでは LeftValue と RightValue はすでに展開済み

        switch (ConditionType)
        {
            case ConditionType.Equal:
                return string.Equals(LeftValue, RightValue, StringComparison.Ordinal);

            case ConditionType.NotEqual:
                return !string.Equals(LeftValue, RightValue, StringComparison.Ordinal);

            case ConditionType.GreaterThan:
                return CompareNumeric((a, b) => a > b);

            case ConditionType.GreaterThanOrEqual:
                return CompareNumeric((a, b) => a >= b);

            case ConditionType.LessThan:
                return CompareNumeric((a, b) => a < b);

            case ConditionType.LessThanOrEqual:
                return CompareNumeric((a, b) => a <= b);

            case ConditionType.Contains:
                return LeftValue.Contains(RightValue, StringComparison.Ordinal);

            case ConditionType.NotContains:
                return !LeftValue.Contains(RightValue, StringComparison.Ordinal);

            case ConditionType.IsEmpty:
                return string.IsNullOrWhiteSpace(LeftValue);

            case ConditionType.IsNotEmpty:
                return !string.IsNullOrWhiteSpace(LeftValue);

            default:
                return false;
        }
    }

    private bool CompareNumeric(Func<double, double, bool> comparison)
    {
        if (!TryParseDouble(LeftValue, out var leftNum))
        {
            LogWarn($"左辺値 '{LeftValue}' は数値ではありません");
            return false;
        }

        if (!TryParseDouble(RightValue, out var rightNum))
        {
            LogWarn($"右辺値 '{RightValue}' は数値ではありません");
            return false;
        }

        return comparison(leftNum, rightNum);
    }

    private bool TryParseDouble(string value, out double result)
    {
        return double.TryParse(value,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture,
            out result);
    }
}
