using Ookii.Dialogs.Wpf;
using RPACore.Actions;
using System.IO;
using System.Windows;

namespace RPAEditor.Dialogs;

public partial class FolderCreateDialog : Window
{
    public FolderCreateAction? Action { get; private set; }

    public FolderCreateDialog()
    {
        InitializeComponent();
    }

    public FolderCreateDialog(FolderCreateAction action) : this()
    {
        // 編集モード
        txtFolderPath.Text = action.FolderPath;
    }

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var folderDialog = new VistaFolderBrowserDialog
        {
            Description = "作成するフォルダの場所を選択してください",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };

        // 現在のパスが入力されている場合、その親フォルダを初期位置に設定
        if (!string.IsNullOrWhiteSpace(txtFolderPath.Text))
        {
            string? parentPath = Path.GetDirectoryName(txtFolderPath.Text);
            if (!string.IsNullOrEmpty(parentPath) && Directory.Exists(parentPath))
            {
                folderDialog.SelectedPath = parentPath;
            }
        }

        if (folderDialog.ShowDialog() == true)
        {
            txtFolderPath.Text = folderDialog.SelectedPath;
        }
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        // 入力検証
        if (string.IsNullOrWhiteSpace(txtFolderPath.Text))
        {
            MessageBox.Show("フォルダパスを指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 無効なパス文字のチェック
        try
        {
            // パスの正規化を試みる（無効なパスの場合は例外が発生）
            string normalizedPath = Path.GetFullPath(txtFolderPath.Text);
        }
        catch (Exception)
        {
            MessageBox.Show("無効なフォルダパスです。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // アクション作成
        Action = new FolderCreateAction
        {
            FolderPath = txtFolderPath.Text
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
