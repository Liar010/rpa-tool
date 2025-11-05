using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using RPACore;
using RPACore.Actions;
using RPAEditor.Services;

namespace RPAEditor.Views;

public partial class EditorPage : UserControl
{
    public event EventHandler? BackRequested;

    private readonly ScriptEngine _scriptEngine;
    private LogViewerWindow? _logViewerWindow;
    private Window? _ownerWindow;
    private readonly RecentFilesManager _recentFilesManager;
    private string? _currentScriptPath;
    private string? _currentScriptFolder;  // スクリプトフォルダのパス
    private bool _isScriptNamed = false;   // スクリプト名が確定しているか

    public EditorPage(ScriptEngine scriptEngine, Window ownerWindow)
    {
        InitializeComponent();

        _scriptEngine = scriptEngine;
        _ownerWindow = ownerWindow;
        _recentFilesManager = new RecentFilesManager();

        // イベントハンドラの登録
        btnBack.Click += BtnBack_Click;
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
        btnAddWebhook.Click += BtnAddWebhook_Click;
        btnAddWaitForImage.Click += BtnAddWaitForImage_Click;
        btnAddScrollUntilImage.Click += BtnAddScrollUntilImage_Click;

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

    private void BtnBack_Click(object sender, RoutedEventArgs e)
    {
        BackRequested?.Invoke(this, EventArgs.Empty);
    }

    public async void LoadScript(string filePath)
    {
        try
        {
            await _scriptEngine.LoadFromFileAsync(filePath);
            _recentFilesManager.AddFile(filePath);
            _currentScriptPath = filePath;

            // スクリプトフォルダを設定（既存スクリプトは名前確定済み）
            _currentScriptFolder = Path.GetDirectoryName(filePath);
            _isScriptNamed = true;

            UpdateActionList();
            txtStatus.Text = $"スクリプトを読み込みました: {Path.GetFileName(filePath)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"スクリプトの読み込みに失敗しました: {ex.Message}", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
            txtStatus.Text = "スクリプトの読み込みに失敗しました";
        }
    }

    private void BtnAddMouseClick_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new MouseActionDialog(_currentScriptPath, _scriptEngine) { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "マウスクリックアクションを追加しました";
        }
    }

    private void BtnAddTextInput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new TextInputDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "テキスト入力アクションを追加しました";
        }
    }

    private void BtnAddKeyPress_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new KeyPressDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "キー入力アクションを追加しました";
        }
    }

    private void BtnAddWait_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new WaitActionDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "待機アクションを追加しました";
        }
    }

    private void BtnAddHotKey_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new HotKeyActionDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "ショートカットキーアクションを追加しました";
        }
    }

    private void BtnAddWindow_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new WindowActionDialog(_scriptEngine.Actions) { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "ウィンドウ操作アクションを追加しました";
        }
    }

    private void BtnAddFileCopy_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileCopyDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイルコピーアクションを追加しました";
        }
    }

    private void BtnAddFileMove_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileMoveDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイル移動アクションを追加しました";
        }
    }

    private void BtnAddFileDelete_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileDeleteDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイル削除アクションを追加しました";
        }
    }

    private void BtnAddFileRename_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileRenameDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイル名変更アクションを追加しました";
        }
    }

    private void BtnAddFolderCreate_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FolderCreateDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "フォルダ作成アクションを追加しました";
        }
    }

    private void BtnAddFileExists_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.FileExistsDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "ファイル/フォルダ存在確認アクションを追加しました";
        }
    }

    private void BtnAddExcelOpen_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelOpenDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excelファイルを開くアクションを追加しました";
        }
    }

    private void BtnAddExcelCell_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelCellDialog(GetActionCount, GetActionAt) { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excelセル操作アクションを追加しました";
        }
    }

    private void BtnAddExcelRange_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelRangeDialog(_scriptEngine) { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excel範囲読み取りアクションを追加しました";
        }
    }

    private void BtnAddExcelSheet_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelSheetDialog(GetActionCount, GetActionAt) { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excelシート操作アクションを追加しました";
        }
    }

    private void BtnAddExcelSaveClose_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.ExcelSaveCloseDialog(GetActionCount, GetActionAt) { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action!);
            UpdateActionList();
            txtStatus.Text = "Excel保存/閉じるアクションを追加しました";
        }
    }

    private void BtnAddWebhook_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.WebhookActionDialog { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "Webhook通知アクションを追加しました";
        }
    }

    private void BtnAddWaitForImage_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new WaitForImageDialog(_currentScriptPath, _scriptEngine) { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "画像待機アクションを追加しました";
        }
    }

    private void BtnAddScrollUntilImage_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new ScrollUntilImageDialog(_currentScriptPath, _scriptEngine) { Owner = _ownerWindow };
        if (dialog.ShowDialog() == true)
        {
            _scriptEngine.AddAction(dialog.Action);
            UpdateActionList();
            txtStatus.Text = "スクロール検索アクションを追加しました";
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

        bool? result = false;
        IAction? newAction = null;

        switch (action)
        {
            case MouseAction mouseAction:
                var mouseDialog = new MouseActionDialog(mouseAction, _currentScriptPath, _scriptEngine) { Owner = _ownerWindow };
                result = mouseDialog.ShowDialog();
                newAction = mouseDialog.Action;
                break;

            case KeyboardAction keyboardAction:
                if (keyboardAction.ActionType == KeyboardActionType.TypeText)
                {
                    var textDialog = new TextInputDialog(keyboardAction) { Owner = _ownerWindow };
                    result = textDialog.ShowDialog();
                    newAction = textDialog.Action;
                }
                else if (keyboardAction.ActionType == KeyboardActionType.PressKey)
                {
                    var keyDialog = new KeyPressDialog(keyboardAction) { Owner = _ownerWindow };
                    result = keyDialog.ShowDialog();
                    newAction = keyDialog.Action;
                }
                else if (keyboardAction.ActionType == KeyboardActionType.HotKey)
                {
                    var hotKeyDialog = new HotKeyActionDialog(keyboardAction) { Owner = _ownerWindow };
                    result = hotKeyDialog.ShowDialog();
                    newAction = hotKeyDialog.Action;
                }
                break;

            case WaitAction waitAction:
                var waitDialog = new WaitActionDialog(waitAction) { Owner = _ownerWindow };
                result = waitDialog.ShowDialog();
                newAction = waitDialog.Action;
                break;

            case WindowAction windowAction:
                var windowDialog = new WindowActionDialog(_scriptEngine.Actions, windowAction, index + 1) { Owner = _ownerWindow };
                result = windowDialog.ShowDialog();
                newAction = windowDialog.Action;
                break;

            case FileCopyAction fileCopyAction:
                var fileCopyDialog = new Dialogs.FileCopyDialog(fileCopyAction) { Owner = _ownerWindow };
                result = fileCopyDialog.ShowDialog();
                newAction = fileCopyDialog.Action;
                break;

            case FileMoveAction fileMoveAction:
                var fileMoveDialog = new Dialogs.FileMoveDialog(fileMoveAction) { Owner = _ownerWindow };
                result = fileMoveDialog.ShowDialog();
                newAction = fileMoveDialog.Action;
                break;

            case FileDeleteAction fileDeleteAction:
                var fileDeleteDialog = new Dialogs.FileDeleteDialog(fileDeleteAction) { Owner = _ownerWindow };
                result = fileDeleteDialog.ShowDialog();
                newAction = fileDeleteDialog.Action;
                break;

            case FileRenameAction fileRenameAction:
                var fileRenameDialog = new Dialogs.FileRenameDialog(fileRenameAction) { Owner = _ownerWindow };
                result = fileRenameDialog.ShowDialog();
                newAction = fileRenameDialog.Action;
                break;

            case FolderCreateAction folderCreateAction:
                var folderCreateDialog = new Dialogs.FolderCreateDialog(folderCreateAction) { Owner = _ownerWindow };
                result = folderCreateDialog.ShowDialog();
                newAction = folderCreateDialog.Action;
                break;

            case FileExistsAction fileExistsAction:
                var fileExistsDialog = new Dialogs.FileExistsDialog(fileExistsAction) { Owner = _ownerWindow };
                result = fileExistsDialog.ShowDialog();
                newAction = fileExistsDialog.Action;
                break;

            case ExcelOpenAction excelOpenAction:
                var excelOpenDialog = new Dialogs.ExcelOpenDialog(excelOpenAction) { Owner = _ownerWindow };
                result = excelOpenDialog.ShowDialog();
                newAction = excelOpenDialog.Action;
                break;

            case ExcelReadCellAction excelReadCellAction:
                var excelReadCellDialog = new Dialogs.ExcelCellDialog(GetActionCount, GetActionAt, excelReadCellAction) { Owner = _ownerWindow };
                result = excelReadCellDialog.ShowDialog();
                newAction = excelReadCellDialog.Action;
                break;

            case ExcelWriteCellAction excelWriteCellAction:
                var excelWriteCellDialog = new Dialogs.ExcelCellDialog(GetActionCount, GetActionAt, excelWriteCellAction) { Owner = _ownerWindow };
                result = excelWriteCellDialog.ShowDialog();
                newAction = excelWriteCellDialog.Action;
                break;

            case ExcelSaveAction excelSaveAction:
                var excelSaveDialog = new Dialogs.ExcelSaveCloseDialog(GetActionCount, GetActionAt, excelSaveAction) { Owner = _ownerWindow };
                result = excelSaveDialog.ShowDialog();
                newAction = excelSaveDialog.Action;
                break;

            case ExcelCloseAction excelCloseAction:
                var excelCloseDialog = new Dialogs.ExcelSaveCloseDialog(GetActionCount, GetActionAt, excelCloseAction) { Owner = _ownerWindow };
                result = excelCloseDialog.ShowDialog();
                newAction = excelCloseDialog.Action;
                break;

            case ExcelAddSheetAction excelAddSheetAction:
                var excelAddSheetDialog = new Dialogs.ExcelSheetDialog(GetActionCount, GetActionAt, excelAddSheetAction) { Owner = _ownerWindow };
                result = excelAddSheetDialog.ShowDialog();
                newAction = excelAddSheetDialog.Action;
                break;

            case ExcelDeleteSheetAction excelDeleteSheetAction:
                var excelDeleteSheetDialog = new Dialogs.ExcelSheetDialog(GetActionCount, GetActionAt, excelDeleteSheetAction) { Owner = _ownerWindow };
                result = excelDeleteSheetDialog.ShowDialog();
                newAction = excelDeleteSheetDialog.Action;
                break;

            case ExcelRenameSheetAction excelRenameSheetAction:
                var excelRenameSheetDialog = new Dialogs.ExcelSheetDialog(GetActionCount, GetActionAt, excelRenameSheetAction) { Owner = _ownerWindow };
                result = excelRenameSheetDialog.ShowDialog();
                newAction = excelRenameSheetDialog.Action;
                break;

            case ExcelCopySheetAction excelCopySheetAction:
                var excelCopySheetDialog = new Dialogs.ExcelSheetDialog(GetActionCount, GetActionAt, excelCopySheetAction) { Owner = _ownerWindow };
                result = excelCopySheetDialog.ShowDialog();
                newAction = excelCopySheetDialog.Action;
                break;

            case ExcelSheetExistsAction excelSheetExistsAction:
                var excelSheetExistsDialog = new Dialogs.ExcelSheetDialog(GetActionCount, GetActionAt, excelSheetExistsAction) { Owner = _ownerWindow };
                result = excelSheetExistsDialog.ShowDialog();
                newAction = excelSheetExistsDialog.Action;
                break;

            case ExcelReadRangeAction excelReadRangeAction:
                var excelReadRangeDialog = new Dialogs.ExcelRangeDialog(excelReadRangeAction, _scriptEngine) { Owner = _ownerWindow };
                result = excelReadRangeDialog.ShowDialog();
                newAction = excelReadRangeDialog.Action;
                break;

            case ExcelWriteRangeAction excelWriteRangeAction:
                var excelWriteRangeDialog = new Dialogs.ExcelRangeDialog(excelWriteRangeAction, _scriptEngine) { Owner = _ownerWindow };
                result = excelWriteRangeDialog.ShowDialog();
                newAction = excelWriteRangeDialog.Action;
                break;

            case WebhookAction webhookAction:
                var webhookDialog = new Dialogs.WebhookActionDialog(webhookAction) { Owner = _ownerWindow };
                result = webhookDialog.ShowDialog();
                newAction = webhookDialog.Action;
                break;

            case WaitForImageAction waitForImageAction:
                var waitForImageDialog = new WaitForImageDialog(waitForImageAction, _currentScriptPath, _scriptEngine) { Owner = _ownerWindow };
                result = waitForImageDialog.ShowDialog();
                newAction = waitForImageDialog.Action;
                break;

            case ScrollUntilImageAction scrollUntilImageAction:
                var scrollUntilImageDialog = new ScrollUntilImageDialog(scrollUntilImageAction, _currentScriptPath, _scriptEngine) { Owner = _ownerWindow };
                result = scrollUntilImageDialog.ShowDialog();
                newAction = scrollUntilImageDialog.Action;
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

        if (success)
        {
            txtStatus.Text = "スクリプト実行完了";
        }
        else
        {
            // エラー概要を1行で表示
            string errorSummary = string.IsNullOrEmpty(_scriptEngine.LastError)
                ? "不明なエラー"
                : _scriptEngine.LastError;
            txtStatus.Text = $"スクリプト実行エラー: {errorSummary}";
        }
    }

    private void BtnStopScript_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show("停止機能は実装中です", "停止", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void BtnLog_Click(object sender, RoutedEventArgs e)
    {
        if (_logViewerWindow != null && _logViewerWindow.IsLoaded)
        {
            _logViewerWindow.Activate();
        }
        else
        {
            _logViewerWindow = new LogViewerWindow { Owner = _ownerWindow };
            _logViewerWindow.Show();
        }
    }

    /// <summary>
    /// 新規スクリプトを作成（仮フォルダ生成）
    /// </summary>
    public void CreateNewScript()
    {
        // Scripts フォルダが存在しなければ作成
        RPACore.ScriptPathManager.EnsureScriptsFolderExists();

        // 新規スクリプト用の一時フォルダを作成
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string folderName = $"新規スクリプト_{timestamp}";
        _currentScriptFolder = Path.Combine(RPACore.ScriptPathManager.ScriptsDirectory, folderName);
        Directory.CreateDirectory(_currentScriptFolder);

        // images フォルダも作成
        Directory.CreateDirectory(Path.Combine(_currentScriptFolder, "images"));

        // スクリプトパスを設定（まだ保存していない）
        _currentScriptPath = Path.Combine(_currentScriptFolder, "script.rpa.json");
        _isScriptNamed = false;

        _scriptEngine.ClearActions();
        UpdateActionList();
        txtStatus.Text = $"新規スクリプトを作成しました: {folderName}（未保存）";
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

        // 新規スクリプトを作成
        CreateNewScript();
    }

    private async void BtnSaveScript_Click(object sender, RoutedEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"Save button clicked. _currentScriptPath={_currentScriptPath}, _isScriptNamed={_isScriptNamed}");

        // スクリプトフォルダが未作成の場合は新規作成
        if (string.IsNullOrEmpty(_currentScriptPath))
        {
            MessageBox.Show("先に「新規」ボタンでスクリプトを作成してください。", "保存エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            // 初回保存時: スクリプト名を確定
            if (!_isScriptNamed)
            {
                string defaultName = Path.GetFileName(_currentScriptFolder ?? "新規スクリプト");
                string? newFolderName = PromptForScriptName(defaultName);

                if (string.IsNullOrWhiteSpace(newFolderName))
                {
                    txtStatus.Text = "保存がキャンセルされました";
                    return;
                }

                // フォルダ名を変更
                newFolderName = newFolderName.Trim();
                string newFolderPath = Path.Combine(RPACore.ScriptPathManager.ScriptsDirectory, newFolderName);

                // 同名フォルダが既に存在するかチェック
                if (Directory.Exists(newFolderPath) && newFolderPath != _currentScriptFolder)
                {
                    MessageBox.Show($"「{newFolderName}」は既に存在します。別の名前を指定してください。",
                        "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // フォルダをリネーム
                if (_currentScriptFolder != null && _currentScriptFolder != newFolderPath)
                {
                    Directory.Move(_currentScriptFolder, newFolderPath);
                    _currentScriptFolder = newFolderPath;
                    _currentScriptPath = Path.Combine(_currentScriptFolder, "script.rpa.json");
                }

                _isScriptNamed = true;
            }

            // スクリプトを保存
            await _scriptEngine.SaveToFileAsync(_currentScriptPath!);
            _recentFilesManager.AddFile(_currentScriptPath!);
            txtStatus.Text = $"スクリプトを保存しました: {Path.GetFileName(_currentScriptFolder)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"スクリプトの保存に失敗しました: {ex.Message}", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
            txtStatus.Text = "スクリプトの保存に失敗しました";
        }
    }

    private async void BtnOpenScript_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "RPA Script (*.rpa.json)|*.rpa.json|All Files (*.*)|*.*",
            DefaultExt = ".rpa.json",
            InitialDirectory = ScriptPathManager.ScriptsDirectory
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                await _scriptEngine.LoadFromFileAsync(dialog.FileName);
                _recentFilesManager.AddFile(dialog.FileName);
                _currentScriptPath = dialog.FileName;

                // スクリプトフォルダを設定（既存スクリプトは名前確定済み）
                _currentScriptFolder = Path.GetDirectoryName(dialog.FileName);
                _isScriptNamed = true;

                UpdateActionList();
                txtStatus.Text = $"スクリプトを読み込みました: {System.IO.Path.GetFileName(dialog.FileName)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"スクリプトの読み込みに失敗しました: {ex.Message}", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                txtStatus.Text = "スクリプトの読み込みに失敗しました";
            }
        }
    }

    private void LstActions_SelectionChanged(object sender, SelectionChangedEventArgs e)
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
                Index = index + 1,
                Action = action
            })
            .ToList();

        lstActions.ItemsSource = null;
        lstActions.ItemsSource = items;
        txtActionCount.Text = $"アクション数: {_scriptEngine.Actions.Count}";
    }

    // Helper methods for dialog constructors
    private int GetActionCount() => _scriptEngine.Actions.Count;
    private IAction GetActionAt(int index) => _scriptEngine.Actions[index];

    /// <summary>
    /// スクリプト名入力ダイアログ（簡易版）
    /// </summary>
    private string? PromptForScriptName(string defaultName)
    {
        var dialog = new Window
        {
            Title = "スクリプト名の設定",
            Width = 400,
            Height = 180,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = _ownerWindow,
            ResizeMode = ResizeMode.NoResize
        };

        var mainStackPanel = new System.Windows.Controls.StackPanel();

        // テキスト入力エリア
        var inputPanel = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        var label = new System.Windows.Controls.Label
        {
            Content = "スクリプト名を入力してください:",
            FontSize = 13
        };
        var textBox = new System.Windows.Controls.TextBox
        {
            Text = defaultName,
            Margin = new Thickness(0, 10, 0, 0),
            Height = 25,
            FontSize = 13
        };

        inputPanel.Children.Add(label);
        inputPanel.Children.Add(textBox);
        mainStackPanel.Children.Add(inputPanel);

        // ボタンエリア
        var buttonPanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(15, 10, 15, 15)
        };

        var okButton = new System.Windows.Controls.Button
        {
            Content = "OK",
            Width = 80,
            Height = 30,
            Margin = new Thickness(5),
            IsDefault = true
        };
        okButton.Click += (s, e) => { dialog.DialogResult = true; dialog.Close(); };

        var cancelButton = new System.Windows.Controls.Button
        {
            Content = "キャンセル",
            Width = 80,
            Height = 30,
            Margin = new Thickness(5),
            IsCancel = true
        };
        cancelButton.Click += (s, e) => { dialog.DialogResult = false; dialog.Close(); };

        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);
        mainStackPanel.Children.Add(buttonPanel);

        dialog.Content = mainStackPanel;

        // ダイアログ表示後にTextBoxにフォーカスと全選択
        dialog.Loaded += (s, e) =>
        {
            textBox.Focus();
            textBox.SelectAll();
        };

        return dialog.ShowDialog() == true ? textBox.Text : null;
    }
}

/// <summary>
/// アクション表示用のビューモデル
/// </summary>
public class ActionListItem
{
    public int Index { get; set; }
    public IAction Action { get; set; } = null!;
    public string DisplayText => $"#{Index}: {Action.Description}";
}
