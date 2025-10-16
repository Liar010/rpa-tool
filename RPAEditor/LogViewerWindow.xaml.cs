using System.Collections.ObjectModel;
using System.Text;
using System.Windows;
using System.Windows.Media;
using RPACore;

namespace RPAEditor;

public partial class LogViewerWindow : Window
{
    private ObservableCollection<LogEntry> _allLogs = new();

    public LogViewerWindow()
    {
        InitializeComponent();

        // 初期ログカウント更新
        UpdateLogCount();

        // Loggerのイベントを購読（InitializeComponent後に実行）
        Logger.Instance.LogAdded += OnLogAdded;

        // ウィンドウが閉じられる時にイベント購読を解除
        Closed += (s, e) => Logger.Instance.LogAdded -= OnLogAdded;
    }

    private void OnLogAdded(object? sender, LogEventArgs e)
    {
        // UIスレッドで実行
        Dispatcher.Invoke(() =>
        {
            // すべてのログに追加
            _allLogs.Add(e.Entry);

            // フィルタに合致すれば表示
            if (ShouldDisplay(e.Entry))
            {
                // TextBoxに追加
                if (txtLogs != null)
                {
                    if (txtLogs.Text.Length > 0)
                        txtLogs.AppendText(Environment.NewLine);
                    txtLogs.AppendText(e.Entry.DisplayText);

                    // 自動スクロール
                    if (chkAutoScroll?.IsChecked == true)
                    {
                        txtLogs.ScrollToEnd();
                    }
                }
            }

            UpdateLogCount();
        });
    }

    private bool ShouldDisplay(LogEntry entry)
    {
        // Null check: XAML初期化中に呼ばれる可能性があるため
        if (chkDebug == null || chkInfo == null || chkWarn == null || chkError == null)
            return true; // デフォルトで表示

        return entry.Level switch
        {
            LogLevel.DEBUG => chkDebug.IsChecked == true,
            LogLevel.INFO => chkInfo.IsChecked == true,
            LogLevel.WARN => chkWarn.IsChecked == true,
            LogLevel.ERROR => chkError.IsChecked == true,
            _ => false
        };
    }

    private void FilterChanged(object sender, RoutedEventArgs e)
    {
        // フィルタが変更されたら表示を更新
        RefreshFilter();
    }

    private void RefreshFilter()
    {
        // Null check: XAML初期化中に呼ばれる可能性があるため
        if (txtLogs == null)
            return;

        // TextBoxをクリアして再構築
        txtLogs.Clear();

        var sb = new StringBuilder();
        int displayCount = 0;

        foreach (var log in _allLogs)
        {
            if (ShouldDisplay(log))
            {
                if (sb.Length > 0)
                    sb.AppendLine();
                sb.Append(log.DisplayText);
                displayCount++;
            }
        }

        txtLogs.Text = sb.ToString();
        UpdateLogCount();

        // 自動スクロール
        if (chkAutoScroll?.IsChecked == true)
        {
            txtLogs.ScrollToEnd();
        }
    }

    private void BtnClear_Click(object sender, RoutedEventArgs e)
    {
        var result = MessageBox.Show(
            "すべてのログをクリアしますか？",
            "確認",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            _allLogs.Clear();
            if (txtLogs != null)
                txtLogs.Clear();
            UpdateLogCount();
        }
    }

    private void UpdateLogCount()
    {
        if (txtLogCount != null)
        {
            int displayCount = txtLogs?.LineCount ?? 0;
            txtLogCount.Text = $"ログ数: {displayCount} / {_allLogs.Count}";
        }
    }
}
