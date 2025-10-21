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
        rbFixedRange.Checked += ReadMode_Changed;
        rbVariableRows.Checked += ReadMode_Changed;
        chkAutoExpandRange.Checked += (s, e) => UpdateControlStates();
        chkAutoExpandRange.Unchecked += (s, e) => UpdateControlStates();
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

        // 可変行数モードか固定範囲モードか
        if (existingAction.ReadVariableRows)
        {
            rbVariableRows.IsChecked = true;
            txtStartColumn.Text = existingAction.StartColumn;
            txtHeaderRow.Text = existingAction.HeaderRow.ToString();
            txtDataStartRow.Text = existingAction.DataStartRow.ToString();
            txtColumnCount.Text = existingAction.ColumnCount.ToString();
        }
        else
        {
            rbFixedRange.IsChecked = true;
            txtRange.Text = existingAction.Range;
        }

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

        chkAutoExpandRange.IsChecked = existingAction.AutoExpandRange;

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

    private void ReadMode_Changed(object sender, RoutedEventArgs e)
    {
        UpdateControlStates();
    }

    private void UpdateControlStates()
    {
        bool isRead = rbRead?.IsChecked == true;
        bool isWrite = rbWrite?.IsChecked == true;

        // 操作選択に応じて読み取りモード選択の表示/非表示
        if (pnlReadMode != null)
            pnlReadMode.Visibility = isRead ? Visibility.Visible : Visibility.Collapsed;

        // 読み取りモードに応じて固定範囲/可変行数パネルの表示/非表示
        bool isFixedRange = rbFixedRange?.IsChecked == true;
        if (isRead)
        {
            if (pnlFixedRange != null)
                pnlFixedRange.Visibility = isFixedRange ? Visibility.Visible : Visibility.Collapsed;
            if (pnlVariableRows != null)
                pnlVariableRows.Visibility = isFixedRange ? Visibility.Collapsed : Visibility.Visible;
        }
        else
        {
            // 書き込み時は常に固定範囲（または開始位置のみ）
            if (pnlFixedRange != null)
                pnlFixedRange.Visibility = Visibility.Visible;
            if (pnlVariableRows != null)
                pnlVariableRows.Visibility = Visibility.Collapsed;
        }

        // 書き込み時はラベルを変更
        if (lblRange != null)
        {
            if (isWrite && chkAutoExpandRange?.IsChecked == true)
            {
                lblRange.Text = "書き込み開始位置 (例: A1):";
            }
            else if (isWrite)
            {
                lblRange.Text = "範囲 (例: A1:B10):";
            }
            else
            {
                lblRange.Text = "範囲 (例: A1:B10):";
            }
        }

        // 操作選択に応じて書き込み値入力欄の表示/非表示
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

        // 読み取りまたは書き込みアクションを作成
        if (rbRead.IsChecked == true)
        {
            // 読み取りアクションのバリデーションと作成
            if (rbVariableRows.IsChecked == true)
            {
                // 可変行数モード
                if (string.IsNullOrWhiteSpace(txtStartColumn.Text))
                {
                    MessageBox.Show("開始列を指定してください（例: A）。", "入力エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!int.TryParse(txtHeaderRow.Text, out int headerRow) || headerRow <= 0)
                {
                    MessageBox.Show("ヘッダー行番号は1以上の整数である必要があります。", "入力エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!int.TryParse(txtDataStartRow.Text, out int dataStartRow) || dataStartRow <= headerRow)
                {
                    MessageBox.Show("データ開始行番号はヘッダー行より後である必要があります。", "入力エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!int.TryParse(txtColumnCount.Text, out int columnCount) || columnCount < 0)
                {
                    MessageBox.Show("列数は0以上の整数である必要があります（0=自動検出）。", "入力エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Action = new ExcelReadRangeAction
                {
                    FilePath = filePath,
                    OpenActionIndex = openActionIndex,
                    SheetName = txtSheetName.Text.Trim(),
                    ReadVariableRows = true,
                    StartColumn = txtStartColumn.Text.Trim(),
                    HeaderRow = headerRow,
                    DataStartRow = dataStartRow,
                    ColumnCount = columnCount
                };
            }
            else
            {
                // 固定範囲モード
                if (string.IsNullOrWhiteSpace(txtRange.Text))
                {
                    MessageBox.Show("範囲を指定してください（例: A1:B10）。", "入力エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!IsValidRangeFormat(txtRange.Text))
                {
                    MessageBox.Show("範囲の形式が正しくありません。\n例: A1:B10, C5:E20", "入力エラー",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Action = new ExcelReadRangeAction
                {
                    FilePath = filePath,
                    OpenActionIndex = openActionIndex,
                    SheetName = txtSheetName.Text.Trim(),
                    ReadVariableRows = false,
                    Range = txtRange.Text.Trim()
                };
            }
        }
        else
        {
            // 書き込みアクション
            if (string.IsNullOrWhiteSpace(txtRange.Text))
            {
                MessageBox.Show("範囲または開始位置を指定してください（例: A1:B10 または A1）。", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 自動拡張モードの場合はセル1つでOK、固定範囲モードの場合は範囲形式をチェック
            bool autoExpand = chkAutoExpandRange.IsChecked == true;
            if (!autoExpand && !IsValidRangeFormat(txtRange.Text) && !IsValidCellAddress(txtRange.Text))
            {
                MessageBox.Show("範囲の形式が正しくありません。\n例: A1:B10, C5:E20", "入力エラー",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

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
                ReadActionIndex = readActionIndex,
                AutoExpandRange = autoExpand
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

    private bool IsValidCellAddress(string cell)
    {
        // 簡易的な検証: "A1" のような形式かチェック
        if (string.IsNullOrWhiteSpace(cell))
            return false;

        return cell.Length > 0 && char.IsLetter(cell[0]);
    }
}
