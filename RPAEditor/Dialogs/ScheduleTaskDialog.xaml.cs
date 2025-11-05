using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using RPACore;
using RPACore.Scheduling;

namespace RPAEditor.Windows;

public partial class ScheduleTaskDialog : Window
{
    public ScheduledTask? Task { get; private set; }

    public ScheduleTaskDialog()
    {
        InitializeComponent();

        // イベントハンドラ登録
        btnBrowseScript.Click += BtnBrowseScript_Click;
        btnOK.Click += BtnOK_Click;
        btnCancel.Click += (s, e) => DialogResult = false;

        rbDaily.Checked += (s, e) => UpdateScheduleTypePanel();
        rbWeekly.Checked += (s, e) => UpdateScheduleTypePanel();
        rbInterval.Checked += (s, e) => UpdateScheduleTypePanel();

        chkNotifyOnCompletion.Checked += (s, e) => pnlNotificationSettings.Visibility = Visibility.Visible;
        chkNotifyOnCompletion.Unchecked += (s, e) => pnlNotificationSettings.Visibility = Visibility.Collapsed;

        // スクリプト一覧を読み込み
        LoadScriptList();

        // デフォルトは毎日
        rbDaily.IsChecked = true;
    }

    public ScheduleTaskDialog(ScheduledTask existingTask) : this()
    {
        Task = existingTask;

        // 既存の値を設定
        txtTaskName.Text = existingTask.Name;

        // スクリプトパスを選択
        SelectScriptPath(existingTask.ScriptPath);

        switch (existingTask.Type)
        {
            case ScheduleType.Daily:
                rbDaily.IsChecked = true;
                txtDailyHour.Text = existingTask.ExecutionTime.Hours.ToString("D2");
                txtDailyMinute.Text = existingTask.ExecutionTime.Minutes.ToString("D2");
                break;

            case ScheduleType.Weekly:
                rbWeekly.IsChecked = true;
                txtWeeklyHour.Text = existingTask.ExecutionTime.Hours.ToString("D2");
                txtWeeklyMinute.Text = existingTask.ExecutionTime.Minutes.ToString("D2");

                if (existingTask.DaysOfWeek != null)
                {
                    chkSunday.IsChecked = existingTask.DaysOfWeek.Contains(DayOfWeek.Sunday);
                    chkMonday.IsChecked = existingTask.DaysOfWeek.Contains(DayOfWeek.Monday);
                    chkTuesday.IsChecked = existingTask.DaysOfWeek.Contains(DayOfWeek.Tuesday);
                    chkWednesday.IsChecked = existingTask.DaysOfWeek.Contains(DayOfWeek.Wednesday);
                    chkThursday.IsChecked = existingTask.DaysOfWeek.Contains(DayOfWeek.Thursday);
                    chkFriday.IsChecked = existingTask.DaysOfWeek.Contains(DayOfWeek.Friday);
                    chkSaturday.IsChecked = existingTask.DaysOfWeek.Contains(DayOfWeek.Saturday);
                }
                break;

            case ScheduleType.Interval:
                rbInterval.IsChecked = true;
                txtIntervalMinutes.Text = existingTask.IntervalMinutes?.ToString() ?? "30";
                break;
        }

        // 通知設定を復元
        chkNotifyOnCompletion.IsChecked = existingTask.NotifyOnCompletion;
        txtNotifyIntervalMinutes.Text = existingTask.NotifyIntervalMinutes.ToString();
        txtNotifyWebhookUrl.Text = existingTask.NotifyWebhookUrl ?? "";

        switch (existingTask.NotifyPolicy)
        {
            case NotificationPolicy.OnError:
                rbNotifyOnError.IsChecked = true;
                break;
            case NotificationPolicy.OnSuccess:
                rbNotifyOnSuccess.IsChecked = true;
                break;
            case NotificationPolicy.Always:
                rbNotifyAlways.IsChecked = true;
                break;
        }
    }

    private void LoadScriptList()
    {
        var scriptItems = new List<ScriptFileItem>();

        try
        {
            ScriptPathManager.EnsureScriptsFolderExists();
            var scriptsDir = ScriptPathManager.ScriptsDirectory;

            if (Directory.Exists(scriptsDir))
            {
                // Scripts フォルダ直下のすべての .rpa.json ファイルを検索（サブフォルダも含む）
                var scriptFiles = Directory.GetFiles(scriptsDir, "*.rpa.json", SearchOption.AllDirectories);

                foreach (var filePath in scriptFiles)
                {
                    var scriptName = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(filePath)); // .rpa.json を除去
                    var relativePath = Path.GetRelativePath(scriptsDir, filePath);

                    scriptItems.Add(new ScriptFileItem
                    {
                        DisplayName = $"{scriptName} ({relativePath})",
                        FullPath = filePath
                    });
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"スクリプト一覧の読み込みに失敗: {ex.Message}");
        }

        // アイテムをコンボボックスに設定
        cmbScriptPath.ItemsSource = scriptItems;

        if (scriptItems.Count > 0)
        {
            cmbScriptPath.SelectedIndex = 0;
        }
    }

    private void SelectScriptPath(string scriptPath)
    {
        if (cmbScriptPath.ItemsSource is List<ScriptFileItem> items)
        {
            var item = items.FirstOrDefault(i => i.FullPath == scriptPath);
            if (item != null)
            {
                cmbScriptPath.SelectedItem = item;
            }
        }
    }

    private void BtnBrowseScript_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "RPAスクリプト (*.rpa.json)|*.rpa.json|すべてのファイル (*.*)|*.*",
            Title = "スクリプトファイルを選択",
            InitialDirectory = ScriptPathManager.ScriptsDirectory
        };

        if (dialog.ShowDialog() == true)
        {
            // 新しいアイテムを追加して選択
            var newItem = new ScriptFileItem
            {
                DisplayName = Path.GetFileName(dialog.FileName),
                FullPath = dialog.FileName
            };

            var items = cmbScriptPath.ItemsSource as List<ScriptFileItem>;
            if (items != null)
            {
                // 既に存在する場合は追加しない
                if (!items.Any(i => i.FullPath == dialog.FileName))
                {
                    items.Add(newItem);
                    cmbScriptPath.ItemsSource = null;
                    cmbScriptPath.ItemsSource = items;
                }
                cmbScriptPath.SelectedItem = items.First(i => i.FullPath == dialog.FileName);
            }
        }
    }

    private void UpdateScheduleTypePanel()
    {
        pnlDaily.Visibility = rbDaily.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        pnlWeekly.Visibility = rbWeekly.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        pnlInterval.Visibility = rbInterval.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        // バリデーション
        if (string.IsNullOrWhiteSpace(txtTaskName.Text))
        {
            MessageBox.Show("タスク名を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (cmbScriptPath.SelectedItem is not ScriptFileItem selectedScript)
        {
            MessageBox.Show("スクリプトファイルを選択してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!File.Exists(selectedScript.FullPath))
        {
            MessageBox.Show("指定されたスクリプトファイルが存在しません", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // タスク作成
        if (Task == null)
        {
            Task = new ScheduledTask();
        }

        Task.Name = txtTaskName.Text;
        Task.ScriptPath = selectedScript.FullPath;

        if (rbDaily.IsChecked == true)
        {
            if (!int.TryParse(txtDailyHour.Text, out int hour) || hour < 0 || hour > 23)
            {
                MessageBox.Show("時刻は00-23の範囲で入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(txtDailyMinute.Text, out int minute) || minute < 0 || minute > 59)
            {
                MessageBox.Show("分は00-59の範囲で入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Task.Type = ScheduleType.Daily;
            Task.ExecutionTime = new TimeSpan(hour, minute, 0);
        }
        else if (rbWeekly.IsChecked == true)
        {
            var selectedDays = new System.Collections.Generic.List<DayOfWeek>();
            if (chkSunday.IsChecked == true) selectedDays.Add(DayOfWeek.Sunday);
            if (chkMonday.IsChecked == true) selectedDays.Add(DayOfWeek.Monday);
            if (chkTuesday.IsChecked == true) selectedDays.Add(DayOfWeek.Tuesday);
            if (chkWednesday.IsChecked == true) selectedDays.Add(DayOfWeek.Wednesday);
            if (chkThursday.IsChecked == true) selectedDays.Add(DayOfWeek.Thursday);
            if (chkFriday.IsChecked == true) selectedDays.Add(DayOfWeek.Friday);
            if (chkSaturday.IsChecked == true) selectedDays.Add(DayOfWeek.Saturday);

            if (selectedDays.Count == 0)
            {
                MessageBox.Show("実行する曜日を少なくとも1つ選択してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(txtWeeklyHour.Text, out int hour) || hour < 0 || hour > 23)
            {
                MessageBox.Show("時刻は00-23の範囲で入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(txtWeeklyMinute.Text, out int minute) || minute < 0 || minute > 59)
            {
                MessageBox.Show("分は00-59の範囲で入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Task.Type = ScheduleType.Weekly;
            Task.ExecutionTime = new TimeSpan(hour, minute, 0);
            Task.DaysOfWeek = selectedDays.ToArray();
        }
        else if (rbInterval.IsChecked == true)
        {
            if (!int.TryParse(txtIntervalMinutes.Text, out int interval) || interval <= 0)
            {
                MessageBox.Show("実行間隔は1以上の数値を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Task.Type = ScheduleType.Interval;
            Task.IntervalMinutes = interval;
        }

        // 通知設定を保存
        Task.NotifyOnCompletion = chkNotifyOnCompletion.IsChecked == true;

        if (Task.NotifyOnCompletion)
        {
            // 通知間隔のバリデーション
            if (!int.TryParse(txtNotifyIntervalMinutes.Text, out int notifyInterval) || notifyInterval < 1)
            {
                MessageBox.Show("通知間隔は1以上の数値を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Task.NotifyIntervalMinutes = notifyInterval;
            Task.NotifyWebhookUrl = string.IsNullOrWhiteSpace(txtNotifyWebhookUrl.Text) ? null : txtNotifyWebhookUrl.Text;

            // 通知ポリシーを設定
            if (rbNotifyOnError.IsChecked == true)
                Task.NotifyPolicy = NotificationPolicy.OnError;
            else if (rbNotifyOnSuccess.IsChecked == true)
                Task.NotifyPolicy = NotificationPolicy.OnSuccess;
            else if (rbNotifyAlways.IsChecked == true)
                Task.NotifyPolicy = NotificationPolicy.Always;
        }

        DialogResult = true;
        Close();
    }
}

/// <summary>
/// スクリプトファイルアイテム（コンボボックス用）
/// </summary>
public class ScriptFileItem
{
    public string DisplayName { get; set; } = "";
    public string FullPath { get; set; } = "";
}
