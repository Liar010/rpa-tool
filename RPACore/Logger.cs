namespace RPACore;

/// <summary>
/// ログレベル
/// </summary>
public enum LogLevel
{
    DEBUG,
    INFO,
    WARN,
    ERROR
}

/// <summary>
/// ログエントリ
/// </summary>
public class LogEntry
{
    /// <summary>タイムスタンプ</summary>
    public DateTime Timestamp { get; set; }

    /// <summary>ログレベル</summary>
    public LogLevel Level { get; set; }

    /// <summary>ログソース（アクション名など）</summary>
    public string Source { get; set; } = string.Empty;

    /// <summary>ログメッセージ</summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>表示用フォーマット</summary>
    public string DisplayText => $"[{Timestamp:HH:mm:ss}] [{Level}] {Source}: {Message}";
}

/// <summary>
/// ログイベント引数
/// </summary>
public class LogEventArgs : EventArgs
{
    public LogEntry Entry { get; }

    public LogEventArgs(LogEntry entry)
    {
        Entry = entry;
    }
}

/// <summary>
/// グローバルロガー（Singleton）
/// </summary>
public class Logger
{
    private static Logger? _instance;
    private static readonly object _lock = new object();

    /// <summary>シングルトンインスタンス</summary>
    public static Logger Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    _instance ??= new Logger();
                }
            }
            return _instance;
        }
    }

    private Logger() { }

    /// <summary>ログ追加イベント</summary>
    public event EventHandler<LogEventArgs>? LogAdded;

    /// <summary>
    /// ログを記録
    /// </summary>
    public void Log(LogLevel level, string source, string message)
    {
        var entry = new LogEntry
        {
            Timestamp = DateTime.Now,
            Level = level,
            Source = source,
            Message = message
        };

        // コンソールにも出力（デバッグ用）
        Console.WriteLine(entry.DisplayText);

        // イベント発火
        LogAdded?.Invoke(this, new LogEventArgs(entry));
    }

    /// <summary>DEBUGレベルログ</summary>
    public void Debug(string source, string message) => Log(LogLevel.DEBUG, source, message);

    /// <summary>INFOレベルログ</summary>
    public void Info(string source, string message) => Log(LogLevel.INFO, source, message);

    /// <summary>WARNレベルログ</summary>
    public void Warn(string source, string message) => Log(LogLevel.WARN, source, message);

    /// <summary>ERRORレベルログ</summary>
    public void Error(string source, string message) => Log(LogLevel.ERROR, source, message);

    /// <summary>すべてのログをクリア（リスナーに通知はしない）</summary>
    public void Clear()
    {
        // 現状はイベントリスナー側で管理するため、ここでは何もしない
        // 将来的にログ履歴を保持する場合はここで実装
    }
}
