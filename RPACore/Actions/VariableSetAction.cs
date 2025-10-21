using System;
using System.Threading.Tasks;

namespace RPACore.Actions;

/// <summary>
/// 変数設定の操作種類
/// </summary>
public enum VariableSetOperationType
{
    Set,       // 直接設定
    Add,       // 数値加算
    Subtract   // 数値減算
}

/// <summary>
/// 変数を設定・操作するアクション
/// </summary>
public class VariableSetAction : ActionBase
{
    /// <summary>変数名</summary>
    public string VariableName { get; set; } = string.Empty;

    /// <summary>設定する値（テンプレート変数や${var}が使用可能）</summary>
    public string Value { get; set; } = string.Empty;

    /// <summary>操作種類</summary>
    public VariableSetOperationType OperationType { get; set; } = VariableSetOperationType.Set;

    public override string Name => "変数設定";

    public override string Description
    {
        get
        {
            var operationSymbol = OperationType switch
            {
                VariableSetOperationType.Set => "=",
                VariableSetOperationType.Add => "+=",
                VariableSetOperationType.Subtract => "-=",
                _ => "="
            };

            return $"変数設定: {VariableName} {operationSymbol} {Value}";
        }
    }

    public override async Task<bool> ExecuteAsync()
    {
        await Task.CompletedTask;

        try
        {
            if (string.IsNullOrWhiteSpace(VariableName))
            {
                LogError("変数名が指定されていません");
                return false;
            }

            if (Context == null)
            {
                LogError("実行コンテキストが利用できません");
                return false;
            }

            LogInfo($"変数操作: {VariableName} {OperationType} {Value}");

            object? result = null;

            switch (OperationType)
            {
                case VariableSetOperationType.Set:
                    // 直接設定（文字列または数値）
                    result = ParseValue(Value);
                    break;

                case VariableSetOperationType.Add:
                    result = PerformArithmetic((a, b) => a + b, "加算");
                    break;

                case VariableSetOperationType.Subtract:
                    result = PerformArithmetic((a, b) => a - b, "減算");
                    break;
            }

            if (result == null)
            {
                LogError("変数操作に失敗しました");
                return false;
            }

            Context.SetVariable(VariableName, result);
            LogInfo($"変数設定完了: {VariableName} = {result}");

            return true;
        }
        catch (Exception ex)
        {
            LogError($"変数設定中にエラーが発生しました: {ex.Message}");
            return false;
        }
    }

    private object? PerformArithmetic(Func<double, double, double> operation, string operationName)
    {
        try
        {
            var currentValue = Context?.GetVariable(VariableName);
            if (currentValue == null)
            {
                LogError($"{operationName}エラー: 変数 '{VariableName}' は未定義です");
                return null;
            }

            if (!TryParseDouble(currentValue.ToString() ?? "0", out var currentNumber))
            {
                LogError($"{operationName}エラー: 変数 '{VariableName}' の値は数値ではありません");
                return null;
            }

            if (!TryParseDouble(Value, out var valueNumber))
            {
                LogError($"{operationName}エラー: 値 '{Value}' は数値ではありません");
                return null;
            }

            var result = operation(currentNumber, valueNumber);
            LogDebug($"{operationName}: {currentNumber} → {result}");

            // 整数の場合はintで返す
            if (Math.Abs(result % 1) < 0.0000001)
            {
                return (int)result;
            }

            return result;
        }
        catch (Exception ex)
        {
            LogError($"{operationName}エラー: {ex.Message}");
            return null;
        }
    }

    private object ParseValue(string value)
    {
        if (TryParseDouble(value, out var number))
        {
            if (Math.Abs(number % 1) < 0.0000001)
            {
                return (int)number;
            }
            return number;
        }

        return value;
    }

    private bool TryParseDouble(string value, out double result)
    {
        return double.TryParse(value,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture,
            out result);
    }
}
