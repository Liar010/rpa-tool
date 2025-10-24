using System.IO;
using System.Windows;
using Microsoft.Win32;
using RPACore;
using RPACore.Actions;

namespace RPAEditor;

public partial class WaitForImageDialog : Window
{
    public WaitForImageAction Action { get; private set; }
    private string? _currentScriptPath;
    private readonly ScriptEngine? _scriptEngine;

    public WaitForImageDialog(string? currentScriptPath = null, ScriptEngine? scriptEngine = null)
    {
        InitializeComponent();
        Action = new WaitForImageAction();
        _currentScriptPath = currentScriptPath;
        _scriptEngine = scriptEngine;

        // イベントハンドラ登録
        btnOK.Click += BtnOK_Click;
        btnSelectImage.Click += BtnSelectImage_Click;
        btnCaptureScreen.Click += BtnCaptureScreen_Click;
        rbWindowNone.Checked += WindowReferenceRadioButton_Checked;
        rbWindowTitle.Checked += WindowReferenceRadioButton_Checked;
        rbWindowLaunchAction.Checked += WindowReferenceRadioButton_Checked;

        UpdateWindowReferenceVisibility();
        PopulateLaunchActions();
    }

    public WaitForImageDialog(WaitForImageAction existingAction, string? currentScriptPath = null, ScriptEngine? scriptEngine = null)
        : this(currentScriptPath, scriptEngine)
    {
        // 既存のアクションの値を設定（編集モード）
        rbModeAppear.IsChecked = existingAction.Mode == WaitForImageMode.Appear;
        rbModeDisappear.IsChecked = existingAction.Mode == WaitForImageMode.Disappear;

        txtTemplatePath.Text = existingAction.TemplateImagePath ?? "";
        txtTimeout.Text = existingAction.TimeoutMs.ToString();
        txtPollingInterval.Text = existingAction.PollingIntervalMs.ToString();
        txtMatchThreshold.Text = existingAction.MatchThreshold.ToString("F2");
        chkMultiScale.IsChecked = existingAction.UseMultiScale;

        // ウィンドウ参照方法の復元
        switch (existingAction.WindowReference)
        {
            case MouseAction.WindowReferenceMethod.None:
                rbWindowNone.IsChecked = true;
                break;
            case MouseAction.WindowReferenceMethod.WindowTitle:
                rbWindowTitle.IsChecked = true;
                txtWindowTitle.Text = existingAction.TargetWindowTitle ?? "";
                break;
            case MouseAction.WindowReferenceMethod.LaunchActionIndex:
                rbWindowLaunchAction.IsChecked = true;
                // ComboBoxから該当アイテムを選択
                for (int i = 0; i < cmbLaunchActions.Items.Count; i++)
                {
                    if (cmbLaunchActions.Items[i] is LaunchActionItem item &&
                        item.ActionIndex == existingAction.TargetLaunchActionIndex)
                    {
                        cmbLaunchActions.SelectedIndex = i;
                        break;
                    }
                }
                break;
        }

        txtSearchX.Text = existingAction.SearchAreaX.ToString();
        txtSearchY.Text = existingAction.SearchAreaY.ToString();
        txtSearchWidth.Text = existingAction.SearchAreaWidth.ToString();
        txtSearchHeight.Text = existingAction.SearchAreaHeight.ToString();

        UpdateWindowReferenceVisibility();
    }

    private void WindowReferenceRadioButton_Checked(object sender, RoutedEventArgs e)
    {
        UpdateWindowReferenceVisibility();
    }

    private void UpdateWindowReferenceVisibility()
    {
        if (pnlWindowTitle == null || pnlLaunchAction == null) return;

        pnlWindowTitle.Visibility = rbWindowTitle.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        pnlLaunchAction.Visibility = rbWindowLaunchAction.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
    }

    private void PopulateLaunchActions()
    {
        cmbLaunchActions.Items.Clear();

        if (_scriptEngine == null) return;

        // WindowActionのLaunch操作のみを抽出
        int index = 1;
        foreach (var action in _scriptEngine.Actions)
        {
            if (action is WindowAction windowAction && windowAction.ActionType == WindowActionType.Launch)
            {
                string displayText = $"#{index}: {windowAction.Description}";
                cmbLaunchActions.Items.Add(new LaunchActionItem
                {
                    ActionIndex = index,
                    DisplayText = displayText
                });
            }
            index++;
        }

        if (cmbLaunchActions.Items.Count > 0)
        {
            cmbLaunchActions.SelectedIndex = 0;
        }
    }

    private class LaunchActionItem
    {
        public int ActionIndex { get; set; }
        public string DisplayText { get; set; } = "";

        public override string ToString() => DisplayText;
    }

    private void BtnSelectImage_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "テンプレート画像を選択",
            Filter = "画像ファイル|*.png;*.jpg;*.jpeg;*.bmp|すべてのファイル|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            txtTemplatePath.Text = dialog.FileName;
        }
    }

    private void BtnCaptureScreen_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentScriptPath))
        {
            MessageBox.Show(
                "画面キャプチャを使用するには、先にスクリプトを保存してください。",
                "スクリプト未保存",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            );
            return;
        }

        // スクリーンキャプチャダイアログを表示
        var captureDialog = new ScreenCaptureDialog(_currentScriptPath);

        // 完了イベントを購読
        captureDialog.CaptureCompleted += (s, args) =>
        {
            if (!string.IsNullOrEmpty(captureDialog.SavedImagePath))
            {
                txtTemplatePath.Text = captureDialog.SavedImagePath;

                // ダイアログが閉じた後にメッセージボックスを表示
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    MessageBox.Show(
                        $"画像を保存しました:\n{captureDialog.SavedImagePath}",
                        "保存完了",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information
                    );
                }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            }
        };

        captureDialog.ShowDialog();
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        // 待機モード
        var mode = rbModeAppear.IsChecked == true ? WaitForImageMode.Appear : WaitForImageMode.Disappear;

        // テンプレート画像
        if (string.IsNullOrEmpty(txtTemplatePath.Text))
        {
            MessageBox.Show("テンプレート画像を選択してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (!File.Exists(txtTemplatePath.Text))
        {
            MessageBox.Show("指定された画像ファイルが見つかりません", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        // タイムアウト
        if (!int.TryParse(txtTimeout.Text, out int timeout) || timeout <= 0)
        {
            MessageBox.Show("タイムアウト時間は正の値で指定してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        // ポーリング間隔
        if (!int.TryParse(txtPollingInterval.Text, out int pollingInterval) || pollingInterval <= 0)
        {
            MessageBox.Show("ポーリング間隔は正の値で指定してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        // 一致率
        if (!double.TryParse(txtMatchThreshold.Text, out double threshold) || threshold < 0 || threshold > 1)
        {
            MessageBox.Show("一致率は0.0～1.0の範囲で指定してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        int searchX = int.TryParse(txtSearchX.Text, out int sx) ? sx : 0;
        int searchY = int.TryParse(txtSearchY.Text, out int sy) ? sy : 0;
        int searchWidth = int.TryParse(txtSearchWidth.Text, out int sw) ? sw : 0;
        int searchHeight = int.TryParse(txtSearchHeight.Text, out int sh) ? sh : 0;

        // ウィンドウ参照方法を決定
        var windowReference = MouseAction.WindowReferenceMethod.None;
        string? targetWindowTitle = null;
        int targetLaunchActionIndex = 0;

        if (rbWindowTitle.IsChecked == true)
        {
            if (string.IsNullOrWhiteSpace(txtWindowTitle.Text))
            {
                MessageBox.Show("ウィンドウタイトルを入力してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            windowReference = MouseAction.WindowReferenceMethod.WindowTitle;
            targetWindowTitle = txtWindowTitle.Text;
        }
        else if (rbWindowLaunchAction.IsChecked == true)
        {
            if (cmbLaunchActions.SelectedItem is LaunchActionItem selectedItem)
            {
                windowReference = MouseAction.WindowReferenceMethod.LaunchActionIndex;
                targetLaunchActionIndex = selectedItem.ActionIndex;
            }
            else
            {
                MessageBox.Show("起動アクションを選択してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        Action = new WaitForImageAction
        {
            Mode = mode,
            TemplateImagePath = txtTemplatePath.Text,
            TimeoutMs = timeout,
            PollingIntervalMs = pollingInterval,
            MatchThreshold = threshold,
            UseMultiScale = chkMultiScale.IsChecked == true,
            WindowReference = windowReference,
            TargetWindowTitle = targetWindowTitle,
            TargetLaunchActionIndex = targetLaunchActionIndex,
            SearchAreaX = searchX,
            SearchAreaY = searchY,
            SearchAreaWidth = searchWidth,
            SearchAreaHeight = searchHeight
        };

        DialogResult = true;
        Close();
    }
}
