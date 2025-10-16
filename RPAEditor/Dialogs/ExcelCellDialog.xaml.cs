using System;
using Microsoft.Win32;
using RPACore.Actions;
using System.Windows;

namespace RPAEditor.Dialogs;

public partial class ExcelCellDialog : Window
{
    public IAction? Action { get; private set; }
    private readonly Func<int>? _getActionCount;
    private readonly Func<int, IAction>? _getActionAt;

    public ExcelCellDialog()
    {
        InitializeComponent();
    }

    public ExcelCellDialog(Func<int> getActionCount, Func<int, IAction> getActionAt) : this()
    {
        _getActionCount = getActionCount;
        _getActionAt = getActionAt;
        PopulateOpenActions();
    }

    public ExcelCellDialog(Func<int> getActionCount, Func<int, IAction> getActionAt, ExcelReadCellAction action)
        : this(getActionCount, getActionAt)
    {
        rbRead.IsChecked = true;

        if (action.OpenActionIndex > 0)
        {
            rbUseOpenAction.IsChecked = true;
            cmbOpenActions.SelectedValue = action.OpenActionIndex;
        }
        else
        {
            rbUseFilePath.IsChecked = true;
            txtFilePath.Text = action.FilePath;
        }

        txtSheetName.Text = action.SheetName;
        txtCellAddress.Text = action.CellAddress;
        UpdateReferenceMethodVisibility();
    }

    public ExcelCellDialog(Func<int> getActionCount, Func<int, IAction> getActionAt, ExcelWriteCellAction action)
        : this(getActionCount, getActionAt)
    {
        rbWrite.IsChecked = true;

        if (action.OpenActionIndex > 0)
        {
            rbUseOpenAction.IsChecked = true;
            cmbOpenActions.SelectedValue = action.OpenActionIndex;
        }
        else
        {
            rbUseFilePath.IsChecked = true;
            txtFilePath.Text = action.FilePath;
        }

        txtSheetName.Text = action.SheetName;
        txtCellAddress.Text = action.CellAddress;

        // 書き込み値の参照方式を設定
        PopulateReadActions();
        if (action.ReadActionIndex > 0)
        {
            rbReadActionValue.IsChecked = true;
            cmbReadActions.SelectedValue = action.ReadActionIndex;
        }
        else
        {
            rbDirectValue.IsChecked = true;
            txtValue.Text = action.Value;
        }

        UpdateReferenceMethodVisibility();
        UpdateValueFieldVisibility();
    }

    private void PopulateOpenActions()
    {
        if (_getActionCount == null || _getActionAt == null) return;

        cmbOpenActions.Items.Clear();
        for (int i = 0; i < _getActionCount(); i++)
        {
            var action = _getActionAt(i);
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

    private void PopulateReadActions()
    {
        if (_getActionCount == null || _getActionAt == null) return;

        cmbReadActions.Items.Clear();
        for (int i = 0; i < _getActionCount(); i++)
        {
            var action = _getActionAt(i);
            if (action is ExcelReadCellAction)
            {
                cmbReadActions.Items.Add(new { Index = i + 1, Display = $"#{i + 1}: セル値を読み取る" });
            }
        }

        cmbReadActions.DisplayMemberPath = "Display";
        cmbReadActions.SelectedValuePath = "Index";

        if (cmbReadActions.Items.Count > 0)
        {
            cmbReadActions.SelectedIndex = 0;
        }
    }

    private void OperationType_Changed(object sender, RoutedEventArgs e)
    {
        UpdateValueFieldVisibility();
        if (rbWrite != null && rbWrite.IsChecked == true)
        {
            PopulateReadActions();
        }
    }

    private void ReferenceMethod_Changed(object sender, RoutedEventArgs e)
    {
        UpdateReferenceMethodVisibility();
    }

    private void ValueMethod_Changed(object sender, RoutedEventArgs e)
    {
        UpdateValueMethodVisibility();
    }

    private void UpdateValueFieldVisibility()
    {
        if (rbWrite == null || lblValueMethod == null || spValueMethod == null ||
            lblDirectValue == null || txtValue == null ||
            lblReadAction == null || cmbReadActions == null)
            return;

        bool isWrite = rbWrite.IsChecked == true;

        lblValueMethod.Visibility = isWrite ? Visibility.Visible : Visibility.Collapsed;
        spValueMethod.Visibility = isWrite ? Visibility.Visible : Visibility.Collapsed;
        lblDirectValue.Visibility = isWrite ? Visibility.Visible : Visibility.Collapsed;
        txtValue.Visibility = isWrite ? Visibility.Visible : Visibility.Collapsed;
        lblReadAction.Visibility = isWrite ? Visibility.Collapsed : Visibility.Collapsed;
        cmbReadActions.Visibility = isWrite ? Visibility.Collapsed : Visibility.Collapsed;

        if (isWrite)
        {
            UpdateValueMethodVisibility();
        }
    }

    private void UpdateValueMethodVisibility()
    {
        if (rbDirectValue == null || lblDirectValue == null || txtValue == null ||
            lblReadAction == null || cmbReadActions == null)
            return;

        bool useDirectValue = rbDirectValue.IsChecked == true;

        lblDirectValue.Visibility = useDirectValue ? Visibility.Visible : Visibility.Collapsed;
        txtValue.Visibility = useDirectValue ? Visibility.Visible : Visibility.Collapsed;
        lblReadAction.Visibility = useDirectValue ? Visibility.Collapsed : Visibility.Visible;
        cmbReadActions.Visibility = useDirectValue ? Visibility.Collapsed : Visibility.Visible;
    }

    private void UpdateReferenceMethodVisibility()
    {
        if (lblFilePath == null || gridFilePath == null || lblOpenAction == null || cmbOpenActions == null)
            return;

        bool useFilePath = rbUseFilePath.IsChecked == true;

        lblFilePath.Visibility = useFilePath ? Visibility.Visible : Visibility.Collapsed;
        gridFilePath.Visibility = useFilePath ? Visibility.Visible : Visibility.Collapsed;

        lblOpenAction.Visibility = useFilePath ? Visibility.Collapsed : Visibility.Visible;
        cmbOpenActions.Visibility = useFilePath ? Visibility.Collapsed : Visibility.Visible;
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
        bool useFilePath = rbUseFilePath.IsChecked == true;

        if (useFilePath && string.IsNullOrWhiteSpace(txtFilePath.Text))
        {
            MessageBox.Show("ファイルパスを指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!useFilePath && cmbOpenActions.SelectedValue == null)
        {
            MessageBox.Show("開くアクションを選択してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(txtCellAddress.Text))
        {
            MessageBox.Show("セルアドレスを指定してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // アクション作成
        if (rbRead.IsChecked == true)
        {
            var action = new ExcelReadCellAction
            {
                SheetName = txtSheetName.Text,
                CellAddress = txtCellAddress.Text
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
            bool useDirectValue = rbDirectValue.IsChecked == true;

            // 書き込み値の入力検証
            if (useDirectValue && string.IsNullOrWhiteSpace(txtValue.Text))
            {
                MessageBox.Show("書き込む値を入力してください。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!useDirectValue && cmbReadActions.SelectedValue == null)
            {
                MessageBox.Show("読み取りアクションを選択してください。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var action = new ExcelWriteCellAction
            {
                SheetName = txtSheetName.Text,
                CellAddress = txtCellAddress.Text
            };

            // ファイル参照方式の設定
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

            // 書き込み値の設定
            if (useDirectValue)
            {
                action.Value = txtValue.Text;
                action.ReadActionIndex = 0;
            }
            else
            {
                action.ReadActionIndex = (int)cmbReadActions.SelectedValue;
                action.Value = string.Empty;
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
