using System;
using System.Threading.Tasks;

namespace RPACore.Actions;

/// <summary>
/// ループ終了条件の種類
/// </summary>
public enum LoopEndConditionType
{
    Always,             // 常にループを抜ける（無限ループ防止用）
    IfEqual,            // 値が等しい場合
    IfNotEqual,         // 値が等しくない場合
    IfGreaterThan,      // 値がより大きい場合
    IfLessThan,         // 値がより小さい場合
    IfEmpty,            // 値が空の場合
    IfNotEmpty,         // 値が空でない場合
    MaxIterations       // 最大繰り返し回数に達した場合
}

/// <summary>
/// ループ終了アクション
/// </summary>
public class LoopEndAction : ActionBase
{
    /// <summary>対応するループ開始アクションの番号</summary>
    public int LoopStartActionIndex { get; set; } = 0;

    /// <summary>終了条件の種類</summary>
    public LoopEndConditionType EndConditionType { get; set; } = LoopEndConditionType.Always;

    /// <summary>比較する左辺値（変数名や固定値）</summary>
    public string LeftValue { get; set; } = string.Empty;

    /// <summary>比較する右辺値（変数名や固定値）</summary>
    public string RightValue { get; set; } = string.Empty;

    /// <summary>最大繰り返し回数（EndConditionType=MaxIterationsの場合）</summary>
    public int MaxIterations { get; set; } = 100;

    /// <summary>現在の繰り返し回数（内部カウンタ）</summary>
    private int _currentIteration = 0;

    public override string Name => "ループ終了";

    public override string Description
    {
        get
        {
            var conditionStr = EndConditionType switch
            {
                LoopEndConditionType.Always => "常に終了",
                LoopEndConditionType.IfEqual => $"{LeftValue} == {RightValue} なら終了",
                LoopEndConditionType.IfNotEqual => $"{LeftValue} != {RightValue} なら終了",
                LoopEndConditionType.IfGreaterThan => $"{LeftValue} > {RightValue} なら終了",
                LoopEndConditionType.IfLessThan => $"{LeftValue} < {RightValue} なら終了",
                LoopEndConditionType.IfEmpty => $"{LeftValue} が空なら終了",
                LoopEndConditionType.IfNotEmpty => $"{LeftValue} が空でないなら終了",
                LoopEndConditionType.MaxIterations => $"{MaxIterations}回繰り返したら終了",
                _ => "?"
            };

            return $"ループ終了 (#{LoopStartActionIndex}へ): {conditionStr}";
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

            // 繰り返し回数をインクリメント
            _currentIteration++;

            // 終了条件を評価
            bool shouldExit = EvaluateEndCondition();

            if (shouldExit)
            {
                LogInfo($"ループを終了します（条件: {EndConditionType}, 繰り返し回数: {_currentIteration}）");
                _currentIteration = 0; // カウンタをリセット
                return true; // 次のアクションに進む
            }
            else
            {
                LogInfo($"ループを継続します（繰り返し回数: {_currentIteration}）");
                // TODO: ScriptEngine でループ開始位置にジャンプする機能を実装
                // 現在は仮実装として true を返す
                return true;
            }
        }
        catch (Exception ex)
        {
            LogError($"ループ終了処理中にエラーが発生しました: {ex.Message}");
            _currentIteration = 0;
            return false;
        }
    }

    private bool EvaluateEndCondition()
    {
        switch (EndConditionType)
        {
            case LoopEndConditionType.Always:
                return true;

            case LoopEndConditionType.IfEqual:
                return string.Equals(LeftValue, RightValue, StringComparison.Ordinal);

            case LoopEndConditionType.IfNotEqual:
                return !string.Equals(LeftValue, RightValue, StringComparison.Ordinal);

            case LoopEndConditionType.IfGreaterThan:
                return CompareNumeric((a, b) => a > b);

            case LoopEndConditionType.IfLessThan:
                return CompareNumeric((a, b) => a < b);

            case LoopEndConditionType.IfEmpty:
                return string.IsNullOrWhiteSpace(LeftValue);

            case LoopEndConditionType.IfNotEmpty:
                return !string.IsNullOrWhiteSpace(LeftValue);

            case LoopEndConditionType.MaxIterations:
                return _currentIteration >= MaxIterations;

            default:
                return true;
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
