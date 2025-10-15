using Microsoft.Win32;
using RPACore.Actions;
using System.IO;
using System.Windows;

namespace RPAEditor.Dialogs;

public partial class FileRenameDialog : Window
{
    public FileRenameAction? Action { get; private set; }

    public FileRenameDialog()
    {
        InitializeComponent();
    }

    public FileRenameDialog(FileRenameAction action) : this()
    {
        // 編集モード
        txtSourcePath.Text = action.SourcePath;
        txtNewName.Text = action.NewName;
    }

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "名前を変更するファイルを選択",
            Filter = "すべてのファイル (*.*)|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            txtSourcePath.Text = dialog.FileName;

            // 現在のファイル名を新しいファイル名の初期値として設定
            if (string.IsNullOrWhiteSpace(txtNewName.Text))
            {
                txtNewName.Text = Path.GetFileName(dialog.FileName);
            }
        }
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        // 入力検証
        if (string.IsNullOrWhiteSpace(txtSourcePath.Text))
        {
            MessageBox.Show("対象ファイルを指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(txtNewName.Text))
        {
            MessageBox.Show("新しいファイル名を指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // パス区切り文字のチェック
        if (txtNewName.Text.Contains(Path.DirectorySeparatorChar) ||
            txtNewName.Text.Contains(Path.AltDirectorySeparatorChar))
        {
            MessageBox.Show("ファイル名にパス区切り文字を含めることはできません。\nファイル名のみを入力してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 無効な文字のチェック
        char[] invalidChars = Path.GetInvalidFileNameChars();
        if (txtNewName.Text.IndexOfAny(invalidChars) >= 0)
        {
            MessageBox.Show("ファイル名に使用できない文字が含まれています。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // アクション作成
        Action = new FileRenameAction
        {
            SourcePath = txtSourcePath.Text,
            NewName = txtNewName.Text
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
