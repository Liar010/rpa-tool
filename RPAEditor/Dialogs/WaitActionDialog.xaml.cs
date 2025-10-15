using System.Windows;
using RPACore.Actions;

namespace RPAEditor;

public partial class WaitActionDialog : Window
{
    public WaitAction Action { get; private set; }

    public WaitActionDialog()
    {
        InitializeComponent();
        Action = new WaitAction();
        btnOK.Click += BtnOK_Click;
    }

    public WaitActionDialog(WaitAction existingAction) : this()
    {
        // 既存のアクションの値を設定（編集モード）
        txtMilliseconds.Text = existingAction.Milliseconds.ToString();
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        if (!int.TryParse(txtMilliseconds.Text, out int milliseconds) || milliseconds <= 0)
        {
            MessageBox.Show("待機時間が無効です（正の整数を入力してください）", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        Action = new WaitAction(milliseconds);
        DialogResult = true;
        Close();
    }
}
