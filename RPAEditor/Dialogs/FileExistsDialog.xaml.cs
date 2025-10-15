using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using RPACore.Actions;
using System.IO;
using System.Windows;

namespace RPAEditor.Dialogs;

public partial class FileExistsDialog : Window
{
    public FileExistsAction? Action { get; private set; }

    public FileExistsDialog()
    {
        InitializeComponent();
    }

    public FileExistsDialog(FileExistsAction action) : this()
    {
        // 編集モード
        txtTargetPath.Text = action.TargetPath;

        if (action.TargetType == FileSystemType.File)
        {
            rbFile.IsChecked = true;
        }
        else
        {
            rbFolder.IsChecked = true;
        }

        chkFailIfNotExists.IsChecked = action.FailIfNotExists;
    }

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        if (rbFile.IsChecked == true)
        {
            // ファイル選択
            var fileDialog = new OpenFileDialog
            {
                Title = "確認対象のファイルを選択",
                Filter = "すべてのファイル (*.*)|*.*"
            };

            if (!string.IsNullOrWhiteSpace(txtTargetPath.Text) && File.Exists(txtTargetPath.Text))
            {
                fileDialog.FileName = txtTargetPath.Text;
            }

            if (fileDialog.ShowDialog() == true)
            {
                txtTargetPath.Text = fileDialog.FileName;
            }
        }
        else
        {
            // フォルダ選択
            var folderDialog = new VistaFolderBrowserDialog
            {
                Description = "確認対象のフォルダを選択してください",
                UseDescriptionForTitle = true
            };

            if (!string.IsNullOrWhiteSpace(txtTargetPath.Text) && Directory.Exists(txtTargetPath.Text))
            {
                folderDialog.SelectedPath = txtTargetPath.Text;
            }

            if (folderDialog.ShowDialog() == true)
            {
                txtTargetPath.Text = folderDialog.SelectedPath;
            }
        }
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        // 入力検証
        if (string.IsNullOrWhiteSpace(txtTargetPath.Text))
        {
            MessageBox.Show("対象パスを指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // アクション作成
        Action = new FileExistsAction
        {
            TargetPath = txtTargetPath.Text,
            TargetType = rbFile.IsChecked == true ? FileSystemType.File : FileSystemType.Folder,
            FailIfNotExists = chkFailIfNotExists.IsChecked == true
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
