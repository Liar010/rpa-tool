using Microsoft.Win32;
using RPACore.Actions;
using System.Windows;

namespace RPAEditor.Dialogs;

public partial class ExcelOpenDialog : Window
{
    public ExcelOpenAction? Action { get; private set; }

    public ExcelOpenDialog()
    {
        InitializeComponent();
    }

    public ExcelOpenDialog(ExcelOpenAction action) : this()
    {
        // 編集モード
        txtFilePath.Text = action.FilePath;
        chkCreateIfNotExists.IsChecked = action.CreateIfNotExists;
    }

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Excelファイルを選択",
            Filter = "Excel ファイル (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|すべてのファイル (*.*)|*.*",
            DefaultExt = ".xlsx"
        };

        if (!string.IsNullOrWhiteSpace(txtFilePath.Text) && System.IO.File.Exists(txtFilePath.Text))
        {
            dialog.FileName = txtFilePath.Text;
        }

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
            MessageBox.Show("ファイルパスを指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // アクション作成
        Action = new ExcelOpenAction
        {
            FilePath = txtFilePath.Text,
            CreateIfNotExists = chkCreateIfNotExists.IsChecked == true
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
