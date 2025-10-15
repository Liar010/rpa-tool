using Microsoft.Win32;
using RPACore.Actions;
using System.IO;
using System.Windows;

namespace RPAEditor.Dialogs;

public partial class FileDeleteDialog : Window
{
    public FileDeleteAction? Action { get; private set; }

    public FileDeleteDialog()
    {
        InitializeComponent();
    }

    public FileDeleteDialog(FileDeleteAction action) : this()
    {
        // 編集モード
        txtFilePath.Text = action.FilePath;
    }

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "削除するファイルを選択",
            Filter = "すべてのファイル (*.*)|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            txtFilePath.Text = dialog.FileName;
        }
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        // 入力検証
        if (string.IsNullOrWhiteSpace(txtFilePath.Text))
        {
            MessageBox.Show("削除するファイルを指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // アクション作成
        Action = new FileDeleteAction
        {
            FilePath = txtFilePath.Text
        };

        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
