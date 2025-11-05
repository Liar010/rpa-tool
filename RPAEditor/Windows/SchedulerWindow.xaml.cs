using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Threading;
using RPACore.Scheduling;

namespace RPAEditor.Windows;

public partial class SchedulerWindow : Window
{
    private readonly SchedulerEngine _schedulerEngine;
    private readonly ObservableCollection<ScheduledTaskViewModel> _taskViewModels;
    private readonly ObservableCollection<ExecutionRecord> _historyRecords;
    private readonly DispatcherTimer _updateTimer;
    private readonly DispatcherTimer _memoryTimer;
    private readonly string _tasksFilePath;

    public SchedulerWindow()
    {
        InitializeComponent();

        _schedulerEngine = new SchedulerEngine();
        _taskViewModels = new ObservableCollection<ScheduledTaskViewModel>();
        _historyRecords = new ObservableCollection<ExecutionRecord>();

        // 保存ファイルパス（exeと同じフォルダ）
        _tasksFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scheduled_tasks.json");

        lstTasks.ItemsSource = _taskViewModels;
        lstHistory.ItemsSource = _historyRecords;

        // イベントハンドラ登録
        btnAddTask.Click += BtnAddTask_Click;
        btnEditTask.Click += BtnEditTask_Click;
        btnDeleteTask.Click += BtnDeleteTask_Click;
        btnStartScheduler.Click += BtnStartScheduler_Click;
        btnStopScheduler.Click += BtnStopScheduler_Click;

        // スケジューラエンジンのイベント
        _schedulerEngine.TaskExecutionStarted += OnTaskExecutionStarted;
        _schedulerEngine.TaskExecutionCompleted += OnTaskExecutionCompleted;

        // UI更新タイマー（1秒ごと）
        _updateTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _updateTimer.Tick += UpdateTimer_Tick;
        _updateTimer.Start();

        // メモリ監視タイマー（5秒ごと）
        _memoryTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
        _memoryTimer.Tick += MemoryTimer_Tick;
        _memoryTimer.Start();

        // 保存済みタスクを読み込み
        _ = LoadTasksAsync();

        UpdateUI();
    }

    private void BtnAddTask_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new ScheduleTaskDialog
        {
            Owner = this
        };

        if (dialog.ShowDialog() == true && dialog.Task != null)
        {
            _schedulerEngine.AddTask(dialog.Task);
            RefreshTaskList();
            _ = SaveTasksAsync(); // 保存
        }
    }

    private void BtnEditTask_Click(object sender, RoutedEventArgs e)
    {
        if (lstTasks.SelectedItem is ScheduledTaskViewModel selectedVM)
        {
            var dialog = new ScheduleTaskDialog(selectedVM.Task)
            {
                Owner = this
            };

            if (dialog.ShowDialog() == true && dialog.Task != null)
            {
                _schedulerEngine.UpdateTask(dialog.Task);
                RefreshTaskList();
                _ = SaveTasksAsync(); // 保存
            }
        }
        else
        {
            MessageBox.Show("編集するタスクを選択してください", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void BtnDeleteTask_Click(object sender, RoutedEventArgs e)
    {
        if (lstTasks.SelectedItem is ScheduledTaskViewModel selectedVM)
        {
            var result = MessageBox.Show(
                $"タスク「{selectedVM.Task.Name}」を削除しますか？",
                "確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _schedulerEngine.RemoveTask(selectedVM.Task.Id);
                RefreshTaskList();
                _ = SaveTasksAsync(); // 保存
            }
        }
        else
        {
            MessageBox.Show("削除するタスクを選択してください", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void BtnStartScheduler_Click(object sender, RoutedEventArgs e)
    {
        _schedulerEngine.Start();
        UpdateUI();
    }

    private void BtnStopScheduler_Click(object sender, RoutedEventArgs e)
    {
        _schedulerEngine.Stop();
        UpdateUI();
    }

    private void OnTaskExecutionStarted(object? sender, TaskExecutionEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            RefreshTaskList();
        });
    }

    private void OnTaskExecutionCompleted(object? sender, TaskExecutionEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            RefreshTaskList();
            RefreshHistory();
        });
    }

    private void UpdateTimer_Tick(object? sender, EventArgs e)
    {
        // 次回実行時刻の表示を更新
        foreach (var vm in _taskViewModels)
        {
            vm.RefreshNextRunText();
        }
    }

    private void MemoryTimer_Tick(object? sender, EventArgs e)
    {
        var process = Process.GetCurrentProcess();
        var memoryMB = process.WorkingSet64 / 1024.0 / 1024.0;
        txtMemoryUsage.Text = $"メモリ使用量: {memoryMB:F1} MB";
    }

    private void RefreshTaskList()
    {
        _taskViewModels.Clear();
        foreach (var task in _schedulerEngine.Tasks)
        {
            _taskViewModels.Add(new ScheduledTaskViewModel(task));
        }
    }

    private void RefreshHistory()
    {
        _historyRecords.Clear();
        foreach (var record in _schedulerEngine.ExecutionHistory.OrderByDescending(r => r.StartTime))
        {
            _historyRecords.Add(record);
        }
    }

    private void UpdateUI()
    {
        if (_schedulerEngine.IsRunning)
        {
            txtSchedulerStatus.Text = "実行中";
            txtSchedulerStatus.Foreground = System.Windows.Media.Brushes.Green;
            btnStartScheduler.IsEnabled = false;
            btnStopScheduler.IsEnabled = true;
        }
        else
        {
            txtSchedulerStatus.Text = "停止中";
            txtSchedulerStatus.Foreground = System.Windows.Media.Brushes.Gray;
            btnStartScheduler.IsEnabled = true;
            btnStopScheduler.IsEnabled = false;
        }
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        _updateTimer?.Stop();
        _memoryTimer?.Stop();
        base.OnClosing(e);
    }

    private async Task LoadTasksAsync()
    {
        await _schedulerEngine.LoadTasksAsync(_tasksFilePath);
        RefreshTaskList();
    }

    private async Task SaveTasksAsync()
    {
        await _schedulerEngine.SaveTasksAsync(_tasksFilePath);
    }
}

/// <summary>
/// タスク表示用ViewModel
/// </summary>
public class ScheduledTaskViewModel : System.ComponentModel.INotifyPropertyChanged
{
    public ScheduledTask Task { get; }

    private string _nextRunText = "";

    public string Name => Task.Name;
    public string EnabledIcon => Task.Enabled ? "✓" : "○";
    public string StatusIcon => Task.Enabled ? "🔔" : "🔕";
    public string ScheduleDescription => Task.GetDescription();

    public string NextRunText
    {
        get => _nextRunText;
        private set
        {
            if (_nextRunText != value)
            {
                _nextRunText = value;
                OnPropertyChanged(nameof(NextRunText));
            }
        }
    }

    public ScheduledTaskViewModel(ScheduledTask task)
    {
        Task = task;
        RefreshNextRunText();
    }

    public void RefreshNextRunText()
    {
        if (Task.NextRun == null)
        {
            NextRunText = "なし";
        }
        else
        {
            var remaining = Task.NextRun.Value - DateTime.Now;
            if (remaining.TotalSeconds < 0)
            {
                NextRunText = Task.NextRun.Value.ToString("MM/dd HH:mm");
            }
            else if (remaining.TotalMinutes < 60)
            {
                NextRunText = $"{(int)remaining.TotalMinutes}分後 ({Task.NextRun.Value:HH:mm})";
            }
            else if (remaining.TotalHours < 24)
            {
                NextRunText = $"{(int)remaining.TotalHours}時間後 ({Task.NextRun.Value:HH:mm})";
            }
            else
            {
                NextRunText = Task.NextRun.Value.ToString("MM/dd HH:mm");
            }
        }
    }

    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
    }
}

/// <summary>
/// Null to Visibility Converter
/// </summary>
public class NullToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return value == null || string.IsNullOrWhiteSpace(value.ToString())
            ? Visibility.Collapsed
            : Visibility.Visible;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
