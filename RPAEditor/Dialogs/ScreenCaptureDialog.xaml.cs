using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using RPACore;

namespace RPAEditor;

public partial class ScreenCaptureDialog : Window
{
    public string? SavedImagePath { get; private set; }
    public event EventHandler? CaptureCompleted;

    private System.Windows.Point _startPoint;
    private bool _isSelecting;
    private readonly string _scriptPath;

    public ScreenCaptureDialog(string scriptPath)
    {
        InitializeComponent();
        _scriptPath = scriptPath;

        // ウィンドウを表示してから少し待機してキャプチャする
        Loaded += async (s, e) =>
        {
            await Task.Delay(100);
            // ウィンドウが透明になるまで待機
        };
    }

    private void Canvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _startPoint = e.GetPosition(canvas);
        _isSelecting = true;
        selectionRect.Visibility = Visibility.Visible;
        Canvas.SetLeft(selectionRect, _startPoint.X);
        Canvas.SetTop(selectionRect, _startPoint.Y);
        selectionRect.Width = 0;
        selectionRect.Height = 0;

        System.Diagnostics.Debug.WriteLine($"MouseDown: {_startPoint.X}, {_startPoint.Y}");
    }

    private void Canvas_MouseMove(object sender, MouseEventArgs e)
    {
        if (!_isSelecting) return;

        var currentPoint = e.GetPosition(canvas);

        double x = Math.Min(_startPoint.X, currentPoint.X);
        double y = Math.Min(_startPoint.Y, currentPoint.Y);
        double width = Math.Abs(currentPoint.X - _startPoint.X);
        double height = Math.Abs(currentPoint.Y - _startPoint.Y);

        Canvas.SetLeft(selectionRect, x);
        Canvas.SetTop(selectionRect, y);
        selectionRect.Width = width;
        selectionRect.Height = height;
    }

    private async void Canvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isSelecting) return;

        _isSelecting = false;

        // 選択範囲を取得（WPF論理ピクセル）
        double x = Canvas.GetLeft(selectionRect);
        double y = Canvas.GetTop(selectionRect);
        double width = selectionRect.Width;
        double height = selectionRect.Height;

        System.Diagnostics.Debug.WriteLine($"MouseUp (logical): x={x}, y={y}, width={width}, height={height}");

        // 最小サイズチェック
        if (width < 10 || height < 10)
        {
            MessageBox.Show("選択範囲が小さすぎます（最小10x10ピクセル）", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            Close();
            return;
        }

        // DPIスケーリングを取得
        var source = PresentationSource.FromVisual(this);
        double dpiScaleX = source?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
        double dpiScaleY = source?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;

        System.Diagnostics.Debug.WriteLine($"DPI Scale: X={dpiScaleX}, Y={dpiScaleY}");

        // ウィンドウの位置を取得（物理ピクセル）
        var windowPosition = this.PointToScreen(new System.Windows.Point(0, 0));

        // Canvas上の座標を物理ピクセルに変換
        int physicalX = (int)(windowPosition.X + x * dpiScaleX);
        int physicalY = (int)(windowPosition.Y + y * dpiScaleY);
        int physicalWidth = (int)(width * dpiScaleX);
        int physicalHeight = (int)(height * dpiScaleY);

        System.Diagnostics.Debug.WriteLine($"Physical pixels: x={physicalX}, y={physicalY}, width={physicalWidth}, height={physicalHeight}");

        // ウィンドウを非表示にしてキャプチャ
        Hide();

        // 少し待機してからキャプチャ（ウィンドウが完全に消えるまで）
        await Task.Delay(200);

        // CaptureAndSaveを実行（完了まで待機）
        await CaptureAndSaveAsync(physicalX, physicalY, physicalWidth, physicalHeight);
    }

    private async Task CaptureAndSaveAsync(int x, int y, int width, int height)
    {
        System.Diagnostics.Debug.WriteLine($"CaptureAndSave called: x={x}, y={y}, w={width}, h={height}");
        System.Diagnostics.Debug.WriteLine($"Script path: {_scriptPath}");

        await Task.Run(() =>
        {
            try
            {
                // スクリーンキャプチャ
                using var bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
                }

                // 保存先パスを生成
                SavedImagePath = ScriptPathManager.GenerateImageFilePath(_scriptPath, "mouse_click");
                System.Diagnostics.Debug.WriteLine($"Image will be saved to: {SavedImagePath}");

                // 画像を保存
                bitmap.Save(SavedImagePath, ImageFormat.Png);
                System.Diagnostics.Debug.WriteLine($"Image saved successfully");
                System.Diagnostics.Debug.WriteLine($"SavedImagePath property is now: {SavedImagePath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in CaptureAndSave: {ex.Message}");
                // UIスレッドでメッセージボックスを表示
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"画像の保存に失敗しました: {ex.Message}\n\n{ex.StackTrace}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                });
                SavedImagePath = null; // エラー時はnullに戻す
            }
        });

        // 保存完了後にイベントを発火してからダイアログを閉じる
        System.Diagnostics.Debug.WriteLine($"About to fire CaptureCompleted event... SavedImagePath={SavedImagePath}");
        CaptureCompleted?.Invoke(this, EventArgs.Empty);
        System.Diagnostics.Debug.WriteLine($"About to close dialog...");
        Close();
        System.Diagnostics.Debug.WriteLine($"Dialog closed");
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            Close();
        }
    }
}
