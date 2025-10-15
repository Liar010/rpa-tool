using System.Runtime.InteropServices;
using System.Windows;
using RPACore.Actions;

namespace RPAEditor;

public partial class MouseActionDialog : Window
{
    public MouseAction Action { get; private set; }

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    public MouseActionDialog()
    {
        InitializeComponent();
        Action = new MouseAction();

        btnOK.Click += BtnOK_Click;
        btnGetMousePos.Click += BtnGetMousePos_Click;
    }

    public MouseActionDialog(MouseAction existingAction) : this()
    {
        // 既存のアクションの値を設定（編集モード）
        txtX.Text = existingAction.X.ToString();
        txtY.Text = existingAction.Y.ToString();
        cmbClickType.SelectedIndex = existingAction.ClickType switch
        {
            MouseClickType.LeftClick => 0,
            MouseClickType.RightClick => 1,
            MouseClickType.DoubleClick => 2,
            MouseClickType.MiddleClick => 3,
            _ => 0
        };
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        if (!int.TryParse(txtX.Text, out int x))
        {
            MessageBox.Show("X座標が無効です", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (!int.TryParse(txtY.Text, out int y))
        {
            MessageBox.Show("Y座標が無効です", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var clickType = cmbClickType.SelectedIndex switch
        {
            0 => MouseClickType.LeftClick,
            1 => MouseClickType.RightClick,
            2 => MouseClickType.DoubleClick,
            3 => MouseClickType.MiddleClick,
            _ => MouseClickType.LeftClick
        };

        Action = new MouseAction(x, y, clickType);
        DialogResult = true;
        Close();
    }

    private async void BtnGetMousePos_Click(object sender, RoutedEventArgs e)
    {
        btnGetMousePos.IsEnabled = false;
        btnGetMousePos.Content = "3秒後にマウス位置を取得...";

        await Task.Delay(3000);

        GetCursorPos(out POINT point);
        txtX.Text = point.X.ToString();
        txtY.Text = point.Y.ToString();

        btnGetMousePos.IsEnabled = true;
        btnGetMousePos.Content = "マウス位置を取得 (3秒後)";
    }
}
