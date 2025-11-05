using System;
using System.Text.Json.Serialization;

namespace RPACore.Scheduling;

/// <summary>
/// スケジュール実行タスク
/// </summary>
public class ScheduledTask
{
    /// <summary>
    /// タスクID（一意識別子）
    /// </summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>
    /// タスク名
    /// </summary>
    public string Name { get; set; } = "";

    /// <summary>
    /// 実行するスクリプトのパス
    /// </summary>
    public string ScriptPath { get; set; } = "";

    /// <summary>
    /// 有効/無効
    /// </summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    /// スケジュールの種類
    /// </summary>
    public ScheduleType Type { get; set; } = ScheduleType.Daily;

    /// <summary>
    /// 実行時刻（Daily/Weekly用）
    /// </summary>
    public TimeSpan ExecutionTime { get; set; } = TimeSpan.Zero;

    /// <summary>
    /// 実行する曜日（Weekly用）
    /// </summary>
    public DayOfWeek[]? DaysOfWeek { get; set; }

    /// <summary>
    /// 実行間隔（分単位、Interval用）
    /// </summary>
    public int? IntervalMinutes { get; set; }

    /// <summary>
    /// 最終実行日時（JSON保存時は除外）
    /// </summary>
    [JsonIgnore]
    public DateTime? LastRun { get; set; }

    /// <summary>
    /// 次回実行予定日時（JSON保存時は除外）
    /// </summary>
    [JsonIgnore]
    public DateTime? NextRun { get; set; }

    /// <summary>
    /// 作成日時
    /// </summary>
    public DateTime CreatedAt { get; set; } = DateTime.Now;

    /// <summary>
    /// 実行完了時に通知する
    /// </summary>
    public bool NotifyOnCompletion { get; set; } = false;

    /// <summary>
    /// 通知ポリシー
    /// </summary>
    public NotificationPolicy NotifyPolicy { get; set; } = NotificationPolicy.OnError;

    /// <summary>
    /// 通知間隔（分）- 連続通知の抑制
    /// </summary>
    public int NotifyIntervalMinutes { get; set; } = 5;

    /// <summary>
    /// 個別のWebhook URL（省略時は設定画面のURLを使用）
    /// </summary>
    public string? NotifyWebhookUrl { get; set; }

    /// <summary>
    /// 最終通知日時（JSON保存時は除外）
    /// </summary>
    [JsonIgnore]
    public DateTime? LastNotificationTime { get; set; }

    /// <summary>
    /// 次回実行日時を計算
    /// </summary>
    public void CalculateNextRun()
    {
        if (!Enabled)
        {
            NextRun = null;
            return;
        }

        var now = DateTime.Now;

        switch (Type)
        {
            case ScheduleType.Daily:
                NextRun = CalculateNextDailyRun(now);
                break;

            case ScheduleType.Weekly:
                NextRun = CalculateNextWeeklyRun(now);
                break;

            case ScheduleType.Interval:
                NextRun = CalculateNextIntervalRun(now);
                break;

            default:
                NextRun = null;
                break;
        }
    }

    private DateTime CalculateNextDailyRun(DateTime now)
    {
        var today = now.Date + ExecutionTime;

        // 今日の実行時刻がまだ来ていなければ今日、過ぎていれば明日
        if (today > now)
        {
            return today;
        }
        else
        {
            return today.AddDays(1);
        }
    }

    private DateTime CalculateNextWeeklyRun(DateTime now)
    {
        if (DaysOfWeek == null || DaysOfWeek.Length == 0)
        {
            // 曜日が指定されていない場合はnull
            return DateTime.MaxValue;
        }

        var today = now.Date + ExecutionTime;

        // 今日から7日間チェック
        for (int i = 0; i < 7; i++)
        {
            var checkDate = today.AddDays(i);
            if (Array.Exists(DaysOfWeek, d => d == checkDate.DayOfWeek))
            {
                if (checkDate > now)
                {
                    return checkDate;
                }
            }
        }

        // 見つからない場合（理論上はありえない）
        return DateTime.MaxValue;
    }

    private DateTime CalculateNextIntervalRun(DateTime now)
    {
        if (IntervalMinutes == null || IntervalMinutes <= 0)
        {
            return DateTime.MaxValue;
        }

        if (LastRun == null)
        {
            // 初回実行は即座に
            return now;
        }

        return LastRun.Value.AddMinutes(IntervalMinutes.Value);
    }

    /// <summary>
    /// タスクの説明文を取得
    /// </summary>
    public string GetDescription()
    {
        switch (Type)
        {
            case ScheduleType.Daily:
                return $"毎日 {ExecutionTime:hh\\:mm}";

            case ScheduleType.Weekly:
                if (DaysOfWeek == null || DaysOfWeek.Length == 0)
                    return "毎週（曜日未設定）";

                var dayNames = string.Join("・", Array.ConvertAll(DaysOfWeek, d => GetDayOfWeekJapanese(d)));
                return $"毎週 {dayNames} {ExecutionTime:hh\\:mm}";

            case ScheduleType.Interval:
                return $"{IntervalMinutes}分ごと";

            default:
                return "不明";
        }
    }

    private static string GetDayOfWeekJapanese(DayOfWeek day)
    {
        return day switch
        {
            DayOfWeek.Sunday => "日",
            DayOfWeek.Monday => "月",
            DayOfWeek.Tuesday => "火",
            DayOfWeek.Wednesday => "水",
            DayOfWeek.Thursday => "木",
            DayOfWeek.Friday => "金",
            DayOfWeek.Saturday => "土",
            _ => "?"
        };
    }
}

/// <summary>
/// スケジュールの種類
/// </summary>
public enum ScheduleType
{
    /// <summary>毎日</summary>
    Daily,

    /// <summary>毎週（曜日指定）</summary>
    Weekly,

    /// <summary>定期実行（分単位）</summary>
    Interval
}

/// <summary>
/// 通知ポリシー
/// </summary>
public enum NotificationPolicy
{
    /// <summary>失敗時のみ通知</summary>
    OnError,

    /// <summary>成功時のみ通知</summary>
    OnSuccess,

    /// <summary>常に通知（成功・失敗両方）</summary>
    Always
}
