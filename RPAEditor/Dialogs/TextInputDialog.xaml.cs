using System.Windows;
using RPACore.Actions;

namespace RPAEditor;

public partial class TextInputDialog : Window
{
    public KeyboardAction Action { get; private set; }

    public TextInputDialog()
    {
        InitializeComponent();
        Action = new KeyboardAction();
        btnOK.Click += BtnOK_Click;
    }

    public TextInputDialog(KeyboardAction existingAction) : this()
    {
        // 既存のアクションの値を設定（編集モード）
        txtText.Text = existingAction.Text;
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtText.Text))
        {
            MessageBox.Show("テキストを入力してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        Action = new KeyboardAction(txtText.Text);
        DialogResult = true;
        Close();
    }
}
