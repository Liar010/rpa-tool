using Microsoft.Win32;
using RPACore.Actions;
using System.Windows;

namespace RPAEditor.Dialogs;

public partial class ExcelSaveCloseDialog : Window
{
    public IAction? Action { get; private set; }
    private readonly MainWindow? _mainWindow;

    public ExcelSaveCloseDialog()
    {
        InitializeComponent();
    }

    public ExcelSaveCloseDialog(MainWindow mainWindow) : this()
    {
        _mainWindow = mainWindow;
        PopulateOpenActions();
    }

    public ExcelSaveCloseDialog(MainWindow mainWindow, ExcelSaveAction action) : this(mainWindow)
    {
        rbSave.IsChecked = true;

        // 参照方式を設定
        if (action.OpenActionIndex > 0)
        {
            rbUseOpenAction.IsChecked = true;
            // ComboBoxから該当するアクションを選択
            for (int i = 0; i < cmbOpenActions.Items.Count; i++)
            {
                var item = cmbOpenActions.Items[i] as dynamic;
                if (item?.Index == action.OpenActionIndex)
                {
                    cmbOpenActions.SelectedIndex = i;
                    break;
                }
            }
        }
        else
        {
            rbUseFilePath.IsChecked = true;
            txtFilePath.Text = action.FilePath;
        }

        txtSaveAsPath.Text = action.SaveAsPath;
        UpdateFieldVisibility();
    }

    public ExcelSaveCloseDialog(MainWindow mainWindow, ExcelCloseAction action) : this(mainWindow)
    {
        rbClose.IsChecked = true;

        // 参照方式を設定
        if (action.OpenActionIndex > 0)
        {
            rbUseOpenAction.IsChecked = true;
            // ComboBoxから該当するアクションを選択
            for (int i = 0; i < cmbOpenActions.Items.Count; i++)
            {
                var item = cmbOpenActions.Items[i] as dynamic;
                if (item?.Index == action.OpenActionIndex)
                {
                    cmbOpenActions.SelectedIndex = i;
                    break;
                }
            }
        }
        else
        {
            rbUseFilePath.IsChecked = true;
            txtFilePath.Text = action.FilePath;
        }

        chkSaveBeforeClose.IsChecked = action.SaveBeforeClose;
        UpdateFieldVisibility();
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

    private void ReferenceMethod_Changed(object sender, RoutedEventArgs e)
    {
        UpdateFieldVisibility();
    }

    private void OperationType_Changed(object sender, RoutedEventArgs e)
    {
        UpdateFieldVisibility();
    }

    private void UpdateFieldVisibility()
    {
        if (lblSaveAsPath != null && gridSaveAsPath != null && chkSaveBeforeClose != null &&
            lblFilePath != null && gridFilePath != null && lblOpenAction != null && cmbOpenActions != null)
        {
            bool isSave = rbSave.IsChecked == true;
            bool useFilePath = rbUseFilePath.IsChecked == true;

            // 参照方式の表示切り替え
            lblFilePath.Visibility = useFilePath ? Visibility.Visible : Visibility.Collapsed;
            gridFilePath.Visibility = useFilePath ? Visibility.Visible : Visibility.Collapsed;
            lblOpenAction.Visibility = useFilePath ? Visibility.Collapsed : Visibility.Visible;
            cmbOpenActions.Visibility = useFilePath ? Visibility.Collapsed : Visibility.Visible;

            // 保存/閉じるの表示切り替え
            lblSaveAsPath.Visibility = isSave ? Visibility.Visible : Visibility.Collapsed;
            gridSaveAsPath.Visibility = isSave ? Visibility.Visible : Visibility.Collapsed;
            chkSaveBeforeClose.Visibility = isSave ? Visibility.Collapsed : Visibility.Visible;
        }
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

    private void BtnBrowseSaveAs_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Title = "保存先を選択",
            Filter = "Excel ファイル (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|すべてのファイル (*.*)|*.*",
            DefaultExt = ".xlsx"
        };

        if (!string.IsNullOrWhiteSpace(txtSaveAsPath.Text))
        {
            dialog.FileName = txtSaveAsPath.Text;
        }

        if (dialog.ShowDialog() == true)
        {
            txtSaveAsPath.Text = dialog.FileName;
        }
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        bool useFilePath = rbUseFilePath.IsChecked == true;

        // 入力検証
        if (useFilePath)
        {
            if (string.IsNullOrWhiteSpace(txtFilePath.Text))
            {
                MessageBox.Show("ファイルパスを指定してください。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        else
        {
            if (cmbOpenActions.SelectedValue == null)
            {
                MessageBox.Show("開くアクションを選択してください。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        // アクション作成
        if (rbSave.IsChecked == true)
        {
            var action = new ExcelSaveAction
            {
                SaveAsPath = txtSaveAsPath.Text
            };

            if (useFilePath)
            {
                action.FilePath = txtFilePath.Text;
                action.OpenActionIndex = 0;
            }
            else
            {
                action.OpenActionIndex = (int)cmbOpenActions.SelectedValue;
                action.FilePath = string.Empty;
            }

            Action = action;
        }
        else
        {
            var action = new ExcelCloseAction
            {
                SaveBeforeClose = chkSaveBeforeClose.IsChecked == true
            };

            if (useFilePath)
            {
                action.FilePath = txtFilePath.Text;
                action.OpenActionIndex = 0;
            }
            else
            {
                action.OpenActionIndex = (int)cmbOpenActions.SelectedValue;
                action.FilePath = string.Empty;
            }

            Action = action;
        }

        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
