using System;

namespace RPACore.Scheduling;

/// <summary>
/// スクリプト実行履歴レコード
/// </summary>
public class ExecutionRecord
{
    /// <summary>
    /// 実行ID
    /// </summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>
    /// タスクID
    /// </summary>
    public string TaskId { get; set; } = "";

    /// <summary>
    /// タスク名（記録時のスナップショット）
    /// </summary>
    public string TaskName { get; set; } = "";

    /// <summary>
    /// スクリプトパス（記録時のスナップショット）
    /// </summary>
    public string ScriptPath { get; set; } = "";

    /// <summary>
    /// 実行開始日時
    /// </summary>
    public DateTime StartTime { get; set; }

    /// <summary>
    /// 実行終了日時
    /// </summary>
    public DateTime? EndTime { get; set; }

    /// <summary>
    /// 実行成功/失敗
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// エラーメッセージ（失敗時）
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// 実行時間（ミリ秒）
    /// </summary>
    public long DurationMs
    {
        get
        {
            if (EndTime == null) return 0;
            return (long)(EndTime.Value - StartTime).TotalMilliseconds;
        }
    }

    /// <summary>
    /// 実行時間の表示用文字列
    /// </summary>
    public string DurationText
    {
        get
        {
            if (EndTime == null) return "実行中...";

            var duration = EndTime.Value - StartTime;
            if (duration.TotalSeconds < 60)
            {
                return $"{duration.TotalSeconds:F1}秒";
            }
            else if (duration.TotalMinutes < 60)
            {
                return $"{duration.TotalMinutes:F1}分";
            }
            else
            {
                return $"{duration.TotalHours:F1}時間";
            }
        }
    }

    /// <summary>
    /// ステータスアイコン
    /// </summary>
    public string StatusIcon => Success ? "✓" : "✗";

    /// <summary>
    /// ステータステキスト
    /// </summary>
    public string StatusText => Success ? "成功" : "失敗";
}
