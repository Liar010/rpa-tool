using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using RPACore;
using RPACore.Actions;

namespace RPAEditor.Dialogs;

public partial class ExcelRangeDialog : Window
{
    public IAction? Action { get; private set; }
    private readonly ScriptEngine? _scriptEngine;
    private readonly bool _isEditMode;

    // 新規作成用コンストラクタ
    public ExcelRangeDialog(ScriptEngine? scriptEngine = null)
    {
        InitializeComponent();
        _scriptEngine = scriptEngine;
        _isEditMode = false;

        // イベントハンドラ登録
        rbRead.Checked += OperationType_Changed;
        rbWrite.Checked += OperationType_Changed;
        rbFilePath.Checked += ReferenceMethod_Changed;
        rbOpenAction.Checked += ReferenceMethod_Changed;
        rbDirectInput.Checked += WriteValueMethod_Changed;
        rbReadReference.Checked += WriteValueMethod_Changed;
        btnBrowse.Click += BtnBrowse_Click;
        btnOK.Click += BtnOK_Click;

        // アクション一覧を設定
        PopulateOpenActions();
        PopulateReadActions();

        // 初期状態設定
        UpdateControlStates();
    }

    // 編集用コンストラクタ（読み取りアクション）
    public ExcelRangeDialog(ExcelReadRangeAction existingAction, ScriptEngine? scriptEngine = null)
        : this(scriptEngine)
    {
        _isEditMode = true;
        rbRead.IsChecked = true;

        // 既存の値を復元
        if (!string.IsNullOrEmpty(existingAction.FilePath))
        {
            rbFilePath.IsChecked = true;
            txtFilePath.Text = existingAction.FilePath;
        }
        else if (existingAction.OpenActionIndex > 0)
        {
            rbOpenAction.IsChecked = true;
            SelectComboBoxItemByTag(cmbOpenAction, existingAction.OpenActionIndex);
        }

        txtSheetName.Text = existingAction.SheetName;
        txtRange.Text = existingAction.Range;

        UpdateControlStates();
    }

    // 編集用コンストラクタ（書き込みアクション）
    public ExcelRangeDialog(ExcelWriteRangeAction existingAction, ScriptEngine? scriptEngine = null)
        : this(scriptEngine)
    {
        _isEditMode = true;
        rbWrite.IsChecked = true;

        // 既存の値を復元
        if (!string.IsNullOrEmpty(existingAction.FilePath))
        {
            rbFilePath.IsChecked = true;
            txtFilePath.Text = existingAction.FilePath;
        }
        else if (existingAction.OpenActionIndex > 0)
        {
            rbOpenAction.IsChecked = true;
            SelectComboBoxItemByTag(cmbOpenAction, existingAction.OpenActionIndex);
        }

        txtSheetName.Text = existingAction.SheetName;
        txtRange.Text = existingAction.Range;

        if (!string.IsNullOrEmpty(existingAction.Value))
        {
            rbDirectInput.IsChecked = true;
            txtValue.Text = existingAction.Value;
        }
        else if (existingAction.ReadActionIndex > 0)
        {
            rbReadReference.IsChecked = true;
            SelectComboBoxItemByTag(cmbReadAction, existingAction.ReadActionIndex);
        }

        UpdateControlStates();
    }

    private void PopulateOpenActions()
    {
        cmbOpenAction.Items.Clear();

        if (_scriptEngine == null)
            return;

        int index = 1;
        foreach (var action in _scriptEngine.Actions)
        {
            if (action is ExcelOpenAction openAction)
            {
                cmbOpenAction.Items.Add(new ComboBoxItem
                {
                    Content = $"#{index}: {openAction.Description}",
                    Tag = index
                });
            }
            index++;
        }

        if (cmbOpenAction.Items.Count > 0)
            cmbOpenAction.SelectedIndex = 0;
    }

    private void PopulateReadActions()
    {
        cmbReadAction.Items.Clear();

        if (_scriptEngine == null)
            return;

        int index = 1;
        foreach (var action in _scriptEngine.Actions)
        {
            if (action is ExcelReadRangeAction readAction)
            {
                cmbReadAction.Items.Add(new ComboBoxItem
                {
                    Content = $"#{index}: {readAction.Description}",
                    Tag = index
                });
            }
            index++;
        }

        if (cmbReadAction.Items.Count > 0)
            cmbReadAction.SelectedIndex = 0;
    }

    private void SelectComboBoxItemByTag(ComboBox comboBox, int tag)
    {
        foreach (ComboBoxItem item in comboBox.Items)
        {
            if ((int)item.Tag == tag)
            {
                comboBox.SelectedItem = item;
                break;
            }
        }
    }

    private void OperationType_Changed(object sender, RoutedEventArgs e)
    {
        UpdateControlStates();
    }

    private void ReferenceMethod_Changed(object sender, RoutedEventArgs e)
    {
        UpdateControlStates();
    }

    private void WriteValueMethod_Changed(object sender, RoutedEventArgs e)
    {
        UpdateControlStates();
    }

    private void UpdateControlStates()
    {
        // 操作選択に応じて書き込み値入力欄の表示/非表示
        bool isWrite = rbWrite?.IsChecked == true;
        if (pnlWriteValue != null)
            pnlWriteValue.Visibility = isWrite ? Visibility.Visible : Visibility.Collapsed;

        // ファイル参照方法に応じて入力欄の有効/無効
        bool useFilePath = rbFilePath?.IsChecked == true;
        if (txtFilePath != null)
            txtFilePath.IsEnabled = useFilePath;
        if (btnBrowse != null)
            btnBrowse.IsEnabled = useFilePath;
        if (cmbOpenAction != null)
            cmbOpenAction.IsEnabled = !useFilePath;

        // 書き込み値入力方法に応じて入力欄の表示/非表示
        if (isWrite)
        {
            bool useDirectInput = rbDirectInput?.IsChecked == true;
            if (lblDirectValue != null)
                lblDirectValue.Visibility = useDirectInput ? Visibility.Visible : Visibility.Collapsed;
            if (txtValue != null)
                txtValue.Visibility = useDirectInput ? Visibility.Visible : Visibility.Collapsed;
            if (lblReadAction != null)
                lblReadAction.Visibility = useDirectInput ? Visibility.Collapsed : Visibility.Visible;
            if (cmbReadAction != null)
                cmbReadAction.Visibility = useDirectInput ? Visibility.Collapsed : Visibility.Visible;
        }
    }

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excelファイル (*.xlsx;*.xls)|*.xlsx;*.xls|すべてのファイル (*.*)|*.*",
            Title = "Excelファイルを選択"
        };

        if (dialog.ShowDialog() == true)
        {
            txtFilePath.Text = dialog.FileName;
        }
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        // 共通バリデーション
        string filePath = string.Empty;
        int openActionIndex = 0;

        if (rbFilePath.IsChecked == true)
        {
            if (string.IsNullOrWhiteSpace(txtFilePath.Text))
            {
                MessageBox.Show("ファイルパスを指定してください。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!File.Exists(txtFilePath.Text))
            {
                MessageBox.Show("指定されたファイルが存在しません。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            filePath = txtFilePath.Text;
        }
        else
        {
            if (cmbOpenAction.SelectedItem == null)
            {
                MessageBox.Show("開くアクションを選択してください。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selectedItem = (ComboBoxItem)cmbOpenAction.SelectedItem;
            openActionIndex = (int)selectedItem.Tag;
        }

        if (string.IsNullOrWhiteSpace(txtRange.Text))
        {
            MessageBox.Show("範囲を指定してください（例: A1:B10）。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 範囲の簡易検証
        if (!IsValidRangeFormat(txtRange.Text))
        {
            MessageBox.Show("範囲の形式が正しくありません。\n例: A1:B10, C5:E20", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 読み取りまたは書き込みアクションを作成
        if (rbRead.IsChecked == true)
        {
            Action = new ExcelReadRangeAction
            {
                FilePath = filePath,
                OpenActionIndex = openActionIndex,
                SheetName = txtSheetName.Text.Trim(),
                Range = txtRange.Text.Trim()
            };
        }
        else
        {
            // 書き込み値のバリデーション
            string value = string.Empty;
            int readActionIndex = 0;

            if (rbDirectInput.IsChecked == true)
            {
                value = txtValue.Text;
            }
            else
            {
                if (cmbReadAction.SelectedItem == null)
                {
                    MessageBox.Show("読み取りアクションを選択してください。", "入力エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var selectedItem = (ComboBoxItem)cmbReadAction.SelectedItem;
                readActionIndex = (int)selectedItem.Tag;
            }

            Action = new ExcelWriteRangeAction
            {
                FilePath = filePath,
                OpenActionIndex = openActionIndex,
                SheetName = txtSheetName.Text.Trim(),
                Range = txtRange.Text.Trim(),
                Value = value,
                ReadActionIndex = readActionIndex
            };
        }

        DialogResult = true;
        Close();
    }

    private bool IsValidRangeFormat(string range)
    {
        // 簡易的な検証: "A1:B10" のような形式かチェック
        if (string.IsNullOrWhiteSpace(range))
            return false;

        var parts = range.Split(':');
        if (parts.Length != 2)
            return false;

        // 各部分がセルアドレスっぽいかチェック（英字で始まる）
        return parts.All(p => !string.IsNullOrWhiteSpace(p) &&
                              p.Length > 0 &&
                              char.IsLetter(p[0]));
    }
}
