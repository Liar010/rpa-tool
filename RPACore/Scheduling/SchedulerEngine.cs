using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace RPACore.Scheduling;

/// <summary>
/// スケジュール実行エンジン
/// </summary>
public class SchedulerEngine
{
    private readonly List<ScheduledTask> _tasks = new();
    private readonly Queue<ExecutionRecord> _executionHistory = new();
    private Timer? _timer;
    private bool _isRunning;
    private readonly object _lockObject = new();

    private const int CHECK_INTERVAL_MS = 10000; // 10秒ごとにチェック
    private const int MAX_HISTORY = 100; // 最大100件の履歴を保持

    /// <summary>
    /// タスクリスト
    /// </summary>
    public IReadOnlyList<ScheduledTask> Tasks
    {
        get
        {
            lock (_lockObject)
            {
                return _tasks.ToList();
            }
        }
    }

    /// <summary>
    /// 実行履歴
    /// </summary>
    public IReadOnlyList<ExecutionRecord> ExecutionHistory
    {
        get
        {
            lock (_lockObject)
            {
                return _executionHistory.ToList();
            }
        }
    }

    /// <summary>
    /// スケジューラが実行中か
    /// </summary>
    public bool IsRunning => _isRunning;

    /// <summary>
    /// タスク実行開始イベント
    /// </summary>
    public event EventHandler<TaskExecutionEventArgs>? TaskExecutionStarted;

    /// <summary>
    /// タスク実行完了イベント
    /// </summary>
    public event EventHandler<TaskExecutionEventArgs>? TaskExecutionCompleted;

    /// <summary>
    /// スケジューラを開始
    /// </summary>
    public void Start()
    {
        if (_isRunning) return;

        _isRunning = true;

        // 全タスクの次回実行時刻を計算
        lock (_lockObject)
        {
            foreach (var task in _tasks)
            {
                task.CalculateNextRun();
            }
        }

        // タイマー開始
        _timer = new Timer(OnTimerElapsed, null, TimeSpan.Zero, TimeSpan.FromMilliseconds(CHECK_INTERVAL_MS));

        Logger.Instance.Info("Scheduler", "スケジューラを開始しました");
    }

    /// <summary>
    /// スケジューラを停止
    /// </summary>
    public void Stop()
    {
        if (!_isRunning) return;

        _isRunning = false;
        _timer?.Dispose();
        _timer = null;

        Logger.Instance.Info("Scheduler", "スケジューラを停止しました");
    }

    /// <summary>
    /// タスクを追加
    /// </summary>
    public void AddTask(ScheduledTask task)
    {
        lock (_lockObject)
        {
            _tasks.Add(task);
            task.CalculateNextRun();
        }

        Logger.Instance.Info("Scheduler", $"スケジュールタスクを追加: {task.Name}");
    }

    /// <summary>
    /// タスクを削除
    /// </summary>
    public void RemoveTask(string taskId)
    {
        lock (_lockObject)
        {
            var task = _tasks.FirstOrDefault(t => t.Id == taskId);
            if (task != null)
            {
                _tasks.Remove(task);
                Logger.Instance.Info("Scheduler", $"スケジュールタスクを削除: {task.Name}");
            }
        }
    }

    /// <summary>
    /// タスクを更新
    /// </summary>
    public void UpdateTask(ScheduledTask task)
    {
        lock (_lockObject)
        {
            var existingTask = _tasks.FirstOrDefault(t => t.Id == task.Id);
            if (existingTask != null)
            {
                _tasks.Remove(existingTask);
                _tasks.Add(task);
                task.CalculateNextRun();
                Logger.Instance.Info("Scheduler", $"スケジュールタスクを更新: {task.Name}");
            }
        }
    }

    /// <summary>
    /// タスクをクリア
    /// </summary>
    public void ClearTasks()
    {
        lock (_lockObject)
        {
            _tasks.Clear();
        }
    }

    /// <summary>
    /// 実行履歴をクリア
    /// </summary>
    public void ClearHistory()
    {
        lock (_lockObject)
        {
            _executionHistory.Clear();
        }
    }

    /// <summary>
    /// タスクをJSONファイルに保存
    /// </summary>
    public async Task SaveTasksAsync(string filePath)
    {
        List<ScheduledTask> tasks;

        lock (_lockObject)
        {
            tasks = _tasks.ToList();
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping // 日本語をエスケープしない
        };

        var json = JsonSerializer.Serialize(tasks, options);
        await File.WriteAllTextAsync(filePath, json, System.Text.Encoding.UTF8);

        Logger.Instance.Info("Scheduler", $"スケジュールタスクを保存しました: {tasks.Count}件");
    }

    /// <summary>
    /// JSONファイルからタスクを読み込み
    /// </summary>
    public async Task LoadTasksAsync(string filePath)
    {
        if (!File.Exists(filePath))
        {
            Logger.Instance.Info("Scheduler", "スケジュールタスクファイルが存在しません");
            return;
        }

        try
        {
            var json = await File.ReadAllTextAsync(filePath, System.Text.Encoding.UTF8);
            var options = new JsonSerializerOptions
            {
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var tasks = JsonSerializer.Deserialize<List<ScheduledTask>>(json, options);

            if (tasks == null)
            {
                Logger.Instance.Warn("Scheduler", "スケジュールタスクの読み込みに失敗しました（nullデータ）");
                return;
            }

            lock (_lockObject)
            {
                _tasks.Clear();

                foreach (var task in tasks)
                {
                    // 次回実行時刻を再計算
                    task.CalculateNextRun();
                    _tasks.Add(task);
                }
            }

            Logger.Instance.Info("Scheduler", $"スケジュールタスクを読み込みました: {tasks.Count}件");
        }
        catch (Exception ex)
        {
            Logger.Instance.Error("Scheduler", $"スケジュールタスクの読み込み中にエラー: {ex.Message}");
        }
    }

    private void OnTimerElapsed(object? state)
    {
        if (!_isRunning) return;

        var now = DateTime.Now;
        ScheduledTask[] tasksToExecute;

        lock (_lockObject)
        {
            // 実行すべきタスクを検索
            tasksToExecute = _tasks
                .Where(t => t.Enabled && t.NextRun != null && t.NextRun <= now)
                .ToArray();
        }

        // タスクを実行（ロック外で実行）
        foreach (var task in tasksToExecute)
        {
            _ = ExecuteTaskAsync(task); // Fire and forget
        }
    }

    private async Task ExecuteTaskAsync(ScheduledTask task)
    {
        var record = new ExecutionRecord
        {
            TaskId = task.Id,
            TaskName = task.Name,
            ScriptPath = task.ScriptPath,
            StartTime = DateTime.Now
        };

        ScriptEngine? scriptEngine = null;

        try
        {
            Logger.Instance.Info("Scheduler", $"スケジュールタスクを実行開始: {task.Name}");
            TaskExecutionStarted?.Invoke(this, new TaskExecutionEventArgs(task, record));

            // スクリプトエンジンで実行
            scriptEngine = new ScriptEngine();
            await scriptEngine.LoadFromFileAsync(task.ScriptPath);
            bool success = await scriptEngine.RunAsync();

            record.EndTime = DateTime.Now;
            record.Success = success;

            if (success)
            {
                Logger.Instance.Info("Scheduler", $"スケジュールタスクを実行完了（成功）: {task.Name}");
            }
            else
            {
                record.ErrorMessage = "スクリプト実行が失敗しました";
                Logger.Instance.Error("Scheduler", $"スケジュールタスクを実行完了（失敗）: {task.Name}");
            }
        }
        catch (Exception ex)
        {
            record.EndTime = DateTime.Now;
            record.Success = false;
            record.ErrorMessage = ex.Message;
            Logger.Instance.Error("Scheduler", $"スケジュールタスク実行中にエラー: {task.Name} - {ex.Message}");
        }
        finally
        {
            // ExecutionContext のクリーンアップ（Excelワークブック等のリソース解放）
            scriptEngine?.Context.Clear();

            // 履歴に追加
            lock (_lockObject)
            {
                _executionHistory.Enqueue(record);

                // 古い履歴を削除
                while (_executionHistory.Count > MAX_HISTORY)
                {
                    _executionHistory.Dequeue();
                }

                // 最終実行時刻を更新
                task.LastRun = record.StartTime;

                // 次回実行時刻を再計算
                task.CalculateNextRun();
            }

            TaskExecutionCompleted?.Invoke(this, new TaskExecutionEventArgs(task, record));

            // 通知機能
            if (ShouldNotify(task, record))
            {
                _ = SendExecutionNotificationAsync(task, record);
            }

            // GC実行（メモリ解放）
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }

    /// <summary>
    /// 通知を送信すべきか判定
    /// </summary>
    private bool ShouldNotify(ScheduledTask task, ExecutionRecord record)
    {
        if (!task.NotifyOnCompletion)
            return false;

        // 通知間隔チェック（レート制限対策）
        if (task.LastNotificationTime != null)
        {
            var elapsed = DateTime.Now - task.LastNotificationTime.Value;
            if (elapsed.TotalMinutes < task.NotifyIntervalMinutes)
            {
                Logger.Instance.Debug("Scheduler",
                    $"通知スキップ（間隔制限）: {task.Name} - 次回通知まで {task.NotifyIntervalMinutes - elapsed.TotalMinutes:F1}分");
                return false;
            }
        }

        // ポリシー別判定
        return task.NotifyPolicy switch
        {
            NotificationPolicy.OnError => !record.Success,
            NotificationPolicy.OnSuccess => record.Success,
            NotificationPolicy.Always => true,
            _ => false
        };
    }

    /// <summary>
    /// 実行通知をWebhookで送信
    /// </summary>
    private async Task SendExecutionNotificationAsync(ScheduledTask task, ExecutionRecord record)
    {
        try
        {
            // Webhook URL を取得（タスク個別設定のみ）
            var webhookUrl = task.NotifyWebhookUrl;

            if (string.IsNullOrWhiteSpace(webhookUrl))
            {
                Logger.Instance.Warn("Scheduler", $"タスク「{task.Name}」のWebhook URLが設定されていないため通知をスキップします");
                return;
            }

            // 通知メッセージを作成
            var message = BuildNotificationMessage(task, record);

            // Webhook送信
            using var client = new System.Net.Http.HttpClient();
            client.Timeout = TimeSpan.FromSeconds(10);

            var payload = new
            {
                text = message
            };

            var json = System.Text.Json.JsonSerializer.Serialize(payload);
            var content = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");

            var response = await client.PostAsync(webhookUrl, content);

            if (response.IsSuccessStatusCode)
            {
                task.LastNotificationTime = DateTime.Now;
                Logger.Instance.Info("Scheduler", $"通知送信成功: {task.Name}");
            }
            else
            {
                Logger.Instance.Warn("Scheduler", $"通知送信失敗（HTTP {response.StatusCode}）: {task.Name}");
            }
        }
        catch (Exception ex)
        {
            Logger.Instance.Error("Scheduler", $"通知送信中にエラー: {task.Name} - {ex.Message}");
        }
    }

    /// <summary>
    /// 通知メッセージを構築
    /// </summary>
    private string BuildNotificationMessage(ScheduledTask task, ExecutionRecord record)
    {
        var statusIcon = record.Success ? "✅" : "❌";
        var statusText = record.Success ? "成功" : "失敗";
        var duration = record.EndTime != null
            ? (record.EndTime.Value - record.StartTime).TotalSeconds.ToString("F1")
            : "不明";

        // スクリプト名は親ディレクトリ名を使用（Scripts/スクリプト名/script.rpa.json の構造）
        var scriptName = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(task.ScriptPath)) ?? "不明";

        var message = $"{statusIcon} スケジュール実行{statusText}\n";
        message += $"タスク名: {task.Name}\n";
        message += $"スクリプト: {scriptName}\n";
        message += $"実行時刻: {record.StartTime:yyyy/MM/dd HH:mm:ss}\n";
        message += $"実行時間: {duration}秒\n";
        message += $"結果: {statusText}";

        if (!record.Success && !string.IsNullOrWhiteSpace(record.ErrorMessage))
        {
            message += $"\nエラー: {record.ErrorMessage}";
        }

        return message;
    }
}

/// <summary>
/// タスク実行イベント引数
/// </summary>
public class TaskExecutionEventArgs : EventArgs
{
    public ScheduledTask Task { get; }
    public ExecutionRecord Record { get; }

    public TaskExecutionEventArgs(ScheduledTask task, ExecutionRecord record)
    {
        Task = task;
        Record = record;
    }
}
