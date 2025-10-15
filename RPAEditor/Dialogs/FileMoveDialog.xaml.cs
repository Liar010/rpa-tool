using Microsoft.Win32;
using RPACore.Actions;
using System.IO;
using System.Windows;
using Ookii.Dialogs.Wpf;

namespace RPAEditor.Dialogs;

public partial class FileMoveDialog : Window
{
    public FileMoveAction? Action { get; private set; }

    public FileMoveDialog()
    {
        InitializeComponent();
    }

    public FileMoveDialog(FileMoveAction action) : this()
    {
        // 編集モード
        txtSourcePath.Text = action.SourcePath;
        txtDestinationPath.Text = action.DestinationPath;
        chkOverwrite.IsChecked = action.Overwrite;
    }

    private void BtnBrowseSource_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "移動元ファイルを選択",
            Filter = "すべてのファイル (*.*)|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            txtSourcePath.Text = dialog.FileName;
        }
    }

    private void BtnBrowseDest_Click(object sender, RoutedEventArgs e)
    {
        // フォルダ選択ダイアログを表示
        var folderDialog = new VistaFolderBrowserDialog
        {
            Description = "移動先フォルダを選択してください",
            UseDescriptionForTitle = true
        };

        if (!string.IsNullOrWhiteSpace(txtDestinationPath.Text) && Directory.Exists(txtDestinationPath.Text))
        {
            folderDialog.SelectedPath = txtDestinationPath.Text;
        }

        if (folderDialog.ShowDialog() == true)
        {
            txtDestinationPath.Text = folderDialog.SelectedPath;
        }
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        // 入力検証
        if (string.IsNullOrWhiteSpace(txtSourcePath.Text))
        {
            MessageBox.Show("移動元ファイルを指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(txtDestinationPath.Text))
        {
            MessageBox.Show("移動先を指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // アクション作成
        Action = new FileMoveAction
        {
            SourcePath = txtSourcePath.Text,
            DestinationPath = txtDestinationPath.Text,
            Overwrite = chkOverwrite.IsChecked == true
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
