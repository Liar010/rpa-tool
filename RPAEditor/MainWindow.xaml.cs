using System.Windows;
using RPACore;
using RPACore.Actions;
using Microsoft.Win32;
using System.Linq;

namespace RPAEditor;

/// <summary>
/// アクション表示用のビューモデル
/// </summary>
public class ActionListItem
{
    public int Index { get; set; }
    public IAction Action { get; set; } = null!;
    public string DisplayText => $"#{Index}: {Action.Description}";
}

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private readonly ScriptEngine _scriptEngine;
    private LogViewerWindow? _logViewerWindow;

    public MainWindow()
    {
        InitializeComponent();
        _scriptEngine = new ScriptEngine();

        // イベントハンドラの登録
        btnAddMouseClick.Click += BtnAddMouseClick_Click;
        btnAddTextInput.Click += BtnAddTextInput_Click;
        btnAddKeyPress.Click += BtnAddKeyPress_Click;
        btnAddHotKey.Click += BtnAddHotKey_Click;
        btnAddWindow.Click += BtnAddWindow_Click;
        btnAddWait.Click += BtnAddWait_Click;
        btnAddFileCopy.Click += BtnAddFileCopy_Click;
        btnAddFileMove.Click += BtnAddFileMove_Click;
        btnAddFileDelete.Click += BtnAddFileDelete_Click;
        btnAddFileRename.Click += BtnAddFileRename_Click;
        btnAddFolderCreate.Click += BtnAddFolderCreate_Click;
        btnAddFileExists.Click += BtnAddFileExists_Click;
        btnAddExcelOpen.Click += BtnAddExcelOpen_Click;
        btnAddExcelCell.Click += BtnAddExcelCell_Click;
        btnAddExcelRange.Click += BtnAddExcelRange_Click;
        btnAddExcelSheet.Click += BtnAddExcelSheet_Click;
        btnAddExcelSaveClose.Click += BtnAddExcelSaveClose_Click;

        btnEditAction.Click += BtnEditAction_Click;
        btnDeleteAction.Click += BtnDeleteAction_Click;
        btnMoveUp.Click += BtnMoveUp_Click;
        btnMoveDown.Click += BtnMoveDown_Click;

        btnRunScript.Click += BtnRunScript_Click;
        btnStopScript.Click += BtnStopScript_Click;
        btnNewScript.Click += BtnNewScript_Click;
        btnSaveScript.Click += BtnSaveScript_Click;
        btnOpenScript.Click += BtnOpenScript_Click;
        btnLog.Click += BtnLog_Click;

        lstActions.SelectionChanged += LstActions_SelectionChanged;

        _scriptEngine.ActionExecuted += ScriptEngine_ActionExecuted;
        _scriptEngine.ScriptCompleted += ScriptEngine_ScriptCompleted;

        UpdateActionList();
    }

    /// <summary>
    /// アクション数を取得
    /// </summary>
    public int GetActionCount() => _scriptEngine.Actions.Count;

    /// <summary>
    /// 指定インデックスのアクションを取得
    /// </summary>
    public IAction GetActionAt(int index) => _scriptEngine.Actions[index];

    private void BtnAddMouseClick_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new MouseActionDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "マウスクリックアクションを追加しました";
        }
    }

    private void BtnAddTextInput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new TextInputDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "テキスト入力アクションを追加しました";
        }
    }

    private void BtnAddKeyPress_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new KeyPressDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "キー入力アクションを追加しました";
        }
    }

    private void BtnAddWait_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new WaitActionDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "待機アクションを追加しました";
        }
    }

    private void BtnAddHotKey_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new HotKeyActionDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "ショートカットキーアクションを追加しました";
        }
    }

    private void BtnAddWindow_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new WindowActionDialog(_scriptEngine.Actions)
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "ウィンドウ操作アクションを追加しました";
        }
    }

    private void BtnAddFileCopy_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileCopyDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイルコピーアクションを追加しました";
        }
    }

    private void BtnAddFileMove_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileMoveDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイル移動アクションを追加しました";
        }
    }

    private void BtnAddFileDelete_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileDeleteDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイル削除アクションを追加しました";
        }
    }

    private void BtnAddFileRename_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileRenameDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイル名変更アクションを追加しました";
        }
    }

    private void BtnAddFolderCreate_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FolderCreateDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "フォルダ作成アクションを追加しました";
        }
    }

    private void BtnAddFileExists_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileExistsDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイル/フォルダ存在確認アクションを追加しました";
        }
    }

    private void BtnAddExcelOpen_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelOpenDialog
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excelファイルを開くアクションを追加しました";
        }
    }

    private void BtnAddExcelCell_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelCellDialog(this)
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excelセル操作アクションを追加しました";
        }
    }

    private void BtnAddExcelRange_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelRangeDialog(_scriptEngine)
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excel範囲読み取りアクションを追加しました";
        }
    }

    private void BtnAddExcelSheet_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelSheetDialog(this)
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excelシート操作アクションを追加しました";
        }
    }

    private void BtnAddExcelSaveClose_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelSaveCloseDialog(this)
        {
            Owner = this
        };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excel保存/閉じるアクションを追加しました";
        }
    }

    private void BtnEditAction_Click(object sender, RoutedEventArgs e)
    {
        if (lstActions.SelectedIndex < 0)
        {
            MessageBox.Show("編集するアクションを選択してください", "編集", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        int index = lstActions.SelectedIndex;
        var action = _scriptEngine.Actions[index];

        // アクションの型に応じて適切なダイアログを表示
        bool? result = false;
        IAction? newAction = null;

        switch (action)
        {
            case MouseAction mouseAction:
                var mouseDialog = new MouseActionDialog(mouseAction) { Owner = this };
                result = mouseDialog.ShowDialog();
                newAction = mouseDialog.Action;
                break;

            case KeyboardAction keyboardAction:
                if (keyboardAction.ActionType == KeyboardActionType.TypeText)
                {
                    var textDialog = new TextInputDialog(keyboardAction) { Owner = this };
                    result = textDialog.ShowDialog();
                    newAction = textDialog.Action;
                }
                else if (keyboardAction.ActionType == KeyboardActionType.PressKey)
                {
                    var keyDialog = new KeyPressDialog(keyboardAction) { Owner = this };
                    result = keyDialog.ShowDialog();
                    newAction = keyDialog.Action;
                }
                else if (keyboardAction.ActionType == KeyboardActionType.HotKey)
                {
                    var hotKeyDialog = new HotKeyActionDialog(keyboardAction) { Owner = this };
                    result = hotKeyDialog.ShowDialog();
                    newAction = hotKeyDialog.Action;
                }
                break;

            case WaitAction waitAction:
                var waitDialog = new WaitActionDialog(waitAction) { Owner = this };
                result = waitDialog.ShowDialog();
                newAction = waitDialog.Action;
                break;

            case WindowAction windowAction:
                var windowDialog = new WindowActionDialog(_scriptEngine.Actions, windowAction, index + 1) { Owner = this };
                result = windowDialog.ShowDialog();
                newAction = windowDialog.Action;
                break;

            case FileCopyAction fileCopyAction:
                var fileCopyDialog = new Dialogs.FileCopyDialog(fileCopyAction) { Owner = this };
                result = fileCopyDialog.ShowDialog();
                newAction = fileCopyDialog.Action;
                break;

            case FileMoveAction fileMoveAction:
                var fileMoveDialog = new Dialogs.FileMoveDialog(fileMoveAction) { Owner = this };
                result = fileMoveDialog.ShowDialog();
                newAction = fileMoveDialog.Action;
                break;

            case FileDeleteAction fileDeleteAction:
                var fileDeleteDialog = new Dialogs.FileDeleteDialog(fileDeleteAction) { Owner = this };
                result = fileDeleteDialog.ShowDialog();
                newAction = fileDeleteDialog.Action;
                break;

            case FileRenameAction fileRenameAction:
                var fileRenameDialog = new Dialogs.FileRenameDialog(fileRenameAction) { Owner = this };
                result = fileRenameDialog.ShowDialog();
                newAction = fileRenameDialog.Action;
                break;

            case FolderCreateAction folderCreateAction:
                var folderCreateDialog = new Dialogs.FolderCreateDialog(folderCreateAction) { Owner = this };
                result = folderCreateDialog.ShowDialog();
                newAction = folderCreateDialog.Action;
                break;

            case FileExistsAction fileExistsAction:
                var fileExistsDialog = new Dialogs.FileExistsDialog(fileExistsAction) { Owner = this };
                result = fileExistsDialog.ShowDialog();
                newAction = fileExistsDialog.Action;
                break;

            case ExcelOpenAction excelOpenAction:
                var excelOpenDialog = new Dialogs.ExcelOpenDialog(excelOpenAction) { Owner = this };
                result = excelOpenDialog.ShowDialog();
                newAction = excelOpenDialog.Action;
                break;

            case ExcelReadCellAction excelReadCellAction:
                var excelReadCellDialog = new Dialogs.ExcelCellDialog(this, excelReadCellAction) { Owner = this };
                result = excelReadCellDialog.ShowDialog();
                newAction = excelReadCellDialog.Action;
                break;

            case ExcelWriteCellAction excelWriteCellAction:
                var excelWriteCellDialog = new Dialogs.ExcelCellDialog(this, excelWriteCellAction) { Owner = this };
                result = excelWriteCellDialog.ShowDialog();
                newAction = excelWriteCellDialog.Action;
                break;

            case ExcelSaveAction excelSaveAction:
                var excelSaveDialog = new Dialogs.ExcelSaveCloseDialog(this, excelSaveAction) { Owner = this };
                result = excelSaveDialog.ShowDialog();
                newAction = excelSaveDialog.Action;
                break;

            case ExcelCloseAction excelCloseAction:
                var excelCloseDialog = new Dialogs.ExcelSaveCloseDialog(this, excelCloseAction) { Owner = this };
                result = excelCloseDialog.ShowDialog();
                newAction = excelCloseDialog.Action;
                break;

            case ExcelAddSheetAction excelAddSheetAction:
                var excelAddSheetDialog = new Dialogs.ExcelSheetDialog(this, excelAddSheetAction) { Owner = this };
                result = excelAddSheetDialog.ShowDialog();
                newAction = excelAddSheetDialog.Action;
                break;

            case ExcelDeleteSheetAction excelDeleteSheetAction:
                var excelDeleteSheetDialog = new Dialogs.ExcelSheetDialog(this, excelDeleteSheetAction) { Owner = this };
                result = excelDeleteSheetDialog.ShowDialog();
                newAction = excelDeleteSheetDialog.Action;
                break;

            case ExcelRenameSheetAction excelRenameSheetAction:
                var excelRenameSheetDialog = new Dialogs.ExcelSheetDialog(this, excelRenameSheetAction) { Owner = this };
                result = excelRenameSheetDialog.ShowDialog();
                newAction = excelRenameSheetDialog.Action;
                break;

            case ExcelCopySheetAction excelCopySheetAction:
                var excelCopySheetDialog = new Dialogs.ExcelSheetDialog(this, excelCopySheetAction) { Owner = this };
                result = excelCopySheetDialog.ShowDialog();
                newAction = excelCopySheetDialog.Action;
                break;

            case ExcelSheetExistsAction excelSheetExistsAction:
                var excelSheetExistsDialog = new Dialogs.ExcelSheetDialog(this, excelSheetExistsAction) { Owner = this };
                result = excelSheetExistsDialog.ShowDialog();
                newAction = excelSheetExistsDialog.Action;
                break;

            case ExcelReadRangeAction excelReadRangeAction:
                var excelReadRangeDialog = new Dialogs.ExcelRangeDialog(excelReadRangeAction, _scriptEngine) { Owner = this };
                result = excelReadRangeDialog.ShowDialog();
                newAction = excelReadRangeDialog.Action;
                break;

            case ExcelWriteRangeAction excelWriteRangeAction:
                var excelWriteRangeDialog = new Dialogs.ExcelRangeDialog(excelWriteRangeAction, _scriptEngine) { Owner = this };
                result = excelWriteRangeDialog.ShowDialog();
                newAction = excelWriteRangeDialog.Action;
                break;

            default:
                MessageBox.Show("このアクションタイプは編集できません", "編集", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
        }

        if (result == true && newAction != null)
        {
            _scriptEngine.Actions[index] = newAction;
            UpdateActionList();
            lstActions.SelectedIndex = index;
            txtStatus.Text = "アクションを編集しました";
        }
    }

    private void BtnDeleteAction_Click(object sender, RoutedEventArgs e)
    {
        if (lstActions.SelectedIndex >= 0)
        {
            var result = MessageBox.Show("選択したアクションを削除しますか？", "確認",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _scriptEngine.Actions.RemoveAt(lstActions.SelectedIndex);
                UpdateActionList();
                txtStatus.Text = "アクションを削除しました";
            }
        }
    }

    private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
    {
        int index = lstActions.SelectedIndex;
        if (index > 0)
        {
            var action = _scriptEngine.Actions[index];
            _scriptEngine.Actions.RemoveAt(index);
            _scriptEngine.Actions.Insert(index - 1, action);
            UpdateActionList();
            lstActions.SelectedIndex = index - 1;
        }
    }

    private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
    {
        int index = lstActions.SelectedIndex;
        if (index >= 0 && index < _scriptEngine.Actions.Count - 1)
        {
            var action = _scriptEngine.Actions[index];
            _scriptEngine.Actions.RemoveAt(index);
            _scriptEngine.Actions.Insert(index + 1, action);
            UpdateActionList();
            lstActions.SelectedIndex = index + 1;
        }
    }

    private async void BtnRunScript_Click(object sender, RoutedEventArgs e)
    {
        if (_scriptEngine.Actions.Count == 0)
        {
            MessageBox.Show("実行するアクションがありません", "実行", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        btnRunScript.IsEnabled = false;
        btnStopScript.IsEnabled = true;
        txtStatus.Text = "スクリプト実行中...";

        bool success = await _scriptEngine.RunAsync();

        btnRunScript.IsEnabled = true;
        btnStopScript.IsEnabled = false;
        txtStatus.Text = success ? "スクリプト実行完了" : "スクリプト実行中にエラーが発生しました";
    }

    private void BtnStopScript_Click(object sender, RoutedEventArgs e)
    {
        // TODO: 実行停止機能を実装
        MessageBox.Show("停止機能は実装中です", "停止", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void BtnLog_Click(object sender, RoutedEventArgs e)
    {
        // ログビューアーが既に開いている場合は前面に表示
        if (_logViewerWindow != null && _logViewerWindow.IsLoaded)
        {
            _logViewerWindow.Activate();
        }
        else
        {
            // 新しくログビューアーを開く
            _logViewerWindow = new LogViewerWindow
            {
                Owner = this
            };
            _logViewerWindow.Show();
        }
    }

    private void BtnNewScript_Click(object sender, RoutedEventArgs e)
    {
        if (_scriptEngine.Actions.Count > 0)
        {
            var result = MessageBox.Show("現在のスクリプトをクリアしますか？", "新規作成",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.No)
                return;
        }

        _scriptEngine.ClearActions();
        UpdateActionList();
        txtStatus.Text = "新規スクリプトを作成しました";
    }

    private async void BtnSaveScript_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Filter = "RPA Script (*.rpa)|*.rpa|All Files (*.*)|*.*",
            DefaultExt = ".rpa"
        };

        if (dialog.ShowDialog() == true)
        {
            await _scriptEngine.SaveToFileAsync(dialog.FileName);
            txtStatus.Text = $"スクリプトを保存しました: {dialog.FileName}";
        }
    }

    private async void BtnOpenScript_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "RPA Script (*.rpa)|*.rpa|All Files (*.*)|*.*",
            DefaultExt = ".rpa"
        };

        if (dialog.ShowDialog() == true)
        {
            await _scriptEngine.LoadFromFileAsync(dialog.FileName);
            UpdateActionList();
            txtStatus.Text = $"スクリプトを読み込みました: {dialog.FileName}";
        }
    }

    private void LstActions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (lstActions.SelectedItem is ActionListItem item)
        {
            txtActionName.Text = item.Action.Name;
            txtActionDescription.Text = item.Action.Description;
        }
        else
        {
            txtActionName.Text = string.Empty;
            txtActionDescription.Text = string.Empty;
        }
    }

    private void ScriptEngine_ActionExecuted(object? sender, ActionExecutedEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            txtExecutionStatus.Text = e.Success ? "✓ 実行成功" : "✗ 実行失敗";
            txtExecutionStatus.Foreground = e.Success ?
                System.Windows.Media.Brushes.Green :
                System.Windows.Media.Brushes.Red;
        });
    }

    private void ScriptEngine_ScriptCompleted(object? sender, ScriptCompletedEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            txtExecutionStatus.Text = "待機中";
            txtExecutionStatus.Foreground = System.Windows.Media.Brushes.Gray;
        });
    }

    private void UpdateActionList()
    {
        var items = _scriptEngine.Actions
            .Select((action, index) => new ActionListItem
            {
                Index = index + 1, // 1-based indexing
                Action = action
            })
            .ToList();

        lstActions.ItemsSource = null;
        lstActions.ItemsSource = items;
        txtActionCount.Text = $"アクション数: {_scriptEngine.Actions.Count}";
    }
}