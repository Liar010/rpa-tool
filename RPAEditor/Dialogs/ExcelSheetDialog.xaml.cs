using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using RPACore.Actions;

namespace RPAEditor.Dialogs;

public partial class ExcelSheetDialog : Window
{
    private readonly MainWindow? _mainWindow;
    public IAction? Action { get; private set; }

    // デフォルトコンストラクタ（新規作成用）
    public ExcelSheetDialog(MainWindow mainWindow)
    {
        InitializeComponent();
        _mainWindow = mainWindow;
        PopulateOpenActions();
        UpdateFieldsVisibility();
    }

    // 編集用コンストラクタ（各アクション型に対応）
    public ExcelSheetDialog(MainWindow mainWindow, IAction existingAction) : this(mainWindow)
    {
        switch (existingAction)
        {
            case ExcelAddSheetAction addAction:
                rbAddSheet.IsChecked = true;
                SetFileReference(addAction.FilePath, addAction.OpenActionIndex);
                txtSheetName.Text = addAction.SheetName;
                break;

            case ExcelDeleteSheetAction deleteAction:
                rbDeleteSheet.IsChecked = true;
                SetFileReference(deleteAction.FilePath, deleteAction.OpenActionIndex);
                txtSheetName.Text = deleteAction.SheetName;
                break;

            case ExcelRenameSheetAction renameAction:
                rbRenameSheet.IsChecked = true;
                SetFileReference(renameAction.FilePath, renameAction.OpenActionIndex);
                txtOldSheetName.Text = renameAction.OldSheetName;
                txtNewSheetName.Text = renameAction.NewSheetName;
                break;

            case ExcelCopySheetAction copyAction:
                rbCopySheet.IsChecked = true;
                SetFileReference(copyAction.FilePath, copyAction.OpenActionIndex);
                txtSourceSheetName.Text = copyAction.SourceSheetName;
                txtDestSheetName.Text = copyAction.DestinationSheetName;
                break;

            case ExcelSheetExistsAction existsAction:
                rbSheetExists.IsChecked = true;
                SetFileReference(existsAction.FilePath, existsAction.OpenActionIndex);
                txtSheetName.Text = existsAction.SheetName;
                break;
        }
    }

    private void SetFileReference(string filePath, int openActionIndex)
    {
        if (openActionIndex > 0)
        {
            rbOpenAction.IsChecked = true;
            SelectOpenActionByIndex(openActionIndex);
        }
        else
        {
            rbFilePath.IsChecked = true;
            txtFilePath.Text = filePath;
        }
    }

    private void SelectOpenActionByIndex(int actionIndex)
    {
        for (int i = 0; i < cmbOpenActions.Items.Count; i++)
        {
            var item = cmbOpenActions.Items[i] as dynamic;
            if (item?.Index == actionIndex)
            {
                cmbOpenActions.SelectedIndex = i;
                break;
            }
        }
    }

    private void PopulateOpenActions()
    {
        if (_mainWindow == null) return;

        cmbOpenActions.Items.Clear();
        for (int i = 0; i < _mainWindow.GetActionCount(); i++)
        {
            var action = _mainWindow.GetActionAt(i);
            if (action is ExcelOpenAction)
            {
                cmbOpenActions.Items.Add(new { Index = i + 1, Display = $"#{i + 1}: Excelファイルを開く" });
            }
        }
        cmbOpenActions.DisplayMemberPath = "Display";
        cmbOpenActions.SelectedValuePath = "Index";
        if (cmbOpenActions.Items.Count > 0)
        {
            cmbOpenActions.SelectedIndex = 0;
        }
    }

    private void OperationType_Changed(object sender, RoutedEventArgs e)
    {
        UpdateFieldsVisibility();
    }

    private void FileRefMethod_Changed(object sender, RoutedEventArgs e)
    {
        UpdateFileRefVisibility();
    }

    private void UpdateFieldsVisibility()
    {
        // Null チェック（XAMLの初期化中に呼ばれる可能性があるため）
        if (rbAddSheet == null || lblSheetName == null) return;

        // すべてのフィールドを一旦非表示
        lblSheetName.Visibility = Visibility.Collapsed;
        txtSheetName.Visibility = Visibility.Collapsed;
        lblOldSheetName.Visibility = Visibility.Collapsed;
        txtOldSheetName.Visibility = Visibility.Collapsed;
        lblNewSheetName.Visibility = Visibility.Collapsed;
        txtNewSheetName.Visibility = Visibility.Collapsed;
        lblSourceSheetName.Visibility = Visibility.Collapsed;
        txtSourceSheetName.Visibility = Visibility.Collapsed;
        lblDestSheetName.Visibility = Visibility.Collapsed;
        txtDestSheetName.Visibility = Visibility.Collapsed;

        // 操作種類に応じて表示するフィールドを選択
        if (rbAddSheet.IsChecked == true)
        {
            lblSheetName.Text = "シート名:";
            lblSheetName.Visibility = Visibility.Visible;
            txtSheetName.Visibility = Visibility.Visible;
            txtDescription.Text = "新しいシートを追加します。";
        }
        else if (rbDeleteSheet.IsChecked == true)
        {
            lblSheetName.Text = "削除するシート名:";
            lblSheetName.Visibility = Visibility.Visible;
            txtSheetName.Visibility = Visibility.Visible;
            txtDescription.Text = "指定したシートを削除します。最後のシートは削除できません。";
        }
        else if (rbRenameSheet.IsChecked == true)
        {
            lblOldSheetName.Visibility = Visibility.Visible;
            txtOldSheetName.Visibility = Visibility.Visible;
            lblNewSheetName.Visibility = Visibility.Visible;
            txtNewSheetName.Visibility = Visibility.Visible;
            txtDescription.Text = "シート名を変更します。";
        }
        else if (rbCopySheet.IsChecked == true)
        {
            lblSourceSheetName.Visibility = Visibility.Visible;
            txtSourceSheetName.Visibility = Visibility.Visible;
            lblDestSheetName.Visibility = Visibility.Visible;
            txtDestSheetName.Visibility = Visibility.Visible;
            txtDescription.Text = "シートをコピーします。内容とフォーマットもコピーされます。";
        }
        else if (rbSheetExists.IsChecked == true)
        {
            lblSheetName.Text = "確認するシート名:";
            lblSheetName.Visibility = Visibility.Visible;
            txtSheetName.Visibility = Visibility.Visible;
            txtDescription.Text = "シートの存在を確認します。結果はアクションのResultプロパティに格納されます。";
        }
    }

    private void UpdateFileRefVisibility()
    {
        if (rbFilePath == null || lblFilePath == null) return;

        bool useFilePath = rbFilePath.IsChecked == true;

        lblFilePath.Visibility = useFilePath ? Visibility.Visible : Visibility.Collapsed;
        dpFilePath.Visibility = useFilePath ? Visibility.Visible : Visibility.Collapsed;
        lblOpenAction.Visibility = useFilePath ? Visibility.Collapsed : Visibility.Visible;
        cmbOpenActions.Visibility = useFilePath ? Visibility.Collapsed : Visibility.Visible;
    }

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
            DefaultExt = ".xlsx"
        };

        if (dialog.ShowDialog() == true)
        {
            txtFilePath.Text = dialog.FileName;
        }
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        // ファイル参照の取得
        string filePath = string.Empty;
        int openActionIndex = 0;

        if (rbFilePath.IsChecked == true)
        {
            filePath = txtFilePath.Text.Trim();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                MessageBox.Show("ファイルパスを入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        else
        {
            if (cmbOpenActions.SelectedItem == null)
            {
                MessageBox.Show("開くアクションを選択してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            var selectedItem = cmbOpenActions.SelectedItem as dynamic;
            openActionIndex = selectedItem!.Index;
        }

        // 操作種類に応じてアクションを作成
        if (rbAddSheet.IsChecked == true)
        {
            var sheetName = txtSheetName.Text.Trim();
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                MessageBox.Show("シート名を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Action = new ExcelAddSheetAction
            {
                FilePath = filePath,
                OpenActionIndex = openActionIndex,
                SheetName = sheetName
            };
        }
        else if (rbDeleteSheet.IsChecked == true)
        {
            var sheetName = txtSheetName.Text.Trim();
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                MessageBox.Show("削除するシート名を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Action = new ExcelDeleteSheetAction
            {
                FilePath = filePath,
                OpenActionIndex = openActionIndex,
                SheetName = sheetName
            };
        }
        else if (rbRenameSheet.IsChecked == true)
        {
            var oldSheetName = txtOldSheetName.Text.Trim();
            var newSheetName = txtNewSheetName.Text.Trim();

            if (string.IsNullOrWhiteSpace(oldSheetName))
            {
                MessageBox.Show("変更前のシート名を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(newSheetName))
            {
                MessageBox.Show("変更後のシート名を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Action = new ExcelRenameSheetAction
            {
                FilePath = filePath,
                OpenActionIndex = openActionIndex,
                OldSheetName = oldSheetName,
                NewSheetName = newSheetName
            };
        }
        else if (rbCopySheet.IsChecked == true)
        {
            var sourceSheetName = txtSourceSheetName.Text.Trim();
            var destSheetName = txtDestSheetName.Text.Trim();

            if (string.IsNullOrWhiteSpace(sourceSheetName))
            {
                MessageBox.Show("コピー元のシート名を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(destSheetName))
            {
                MessageBox.Show("コピー先のシート名を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Action = new ExcelCopySheetAction
            {
                FilePath = filePath,
                OpenActionIndex = openActionIndex,
                SourceSheetName = sourceSheetName,
                DestinationSheetName = destSheetName
            };
        }
        else if (rbSheetExists.IsChecked == true)
        {
            var sheetName = txtSheetName.Text.Trim();
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                MessageBox.Show("確認するシート名を入力してください", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Action = new ExcelSheetExistsAction
            {
                FilePath = filePath,
                OpenActionIndex = openActionIndex,
                SheetName = sheetName
            };
        }

        DialogResult = true;
        Close();
    }
}
