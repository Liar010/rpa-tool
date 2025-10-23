using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using RPACore;
using RPACore.Actions;

namespace RPAEditor;

public partial class MouseActionDialog : Window
{
    public MouseAction Action { get; private set; }
    private string? _currentScriptPath;
    private readonly ScriptEngine? _scriptEngine;

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    public MouseActionDialog(string? currentScriptPath = null, ScriptEngine? scriptEngine = null)
    {
        InitializeComponent();
        Action = new MouseAction();
        _currentScriptPath = currentScriptPath;
        _scriptEngine = scriptEngine;

        // イベントハンドラ登録
        btnOK.Click += BtnOK_Click;
        btnGetMousePos.Click += BtnGetMousePos_Click;
        btnSelectImage.Click += BtnSelectImage_Click;
        btnCaptureScreen.Click += BtnCaptureScreen_Click;
        rbCoordinate.Checked += RadioButton_Checked;
        rbImageMatch.Checked += RadioButton_Checked;
        rbWindowNone.Checked += WindowReferenceRadioButton_Checked;
        rbWindowTitle.Checked += WindowReferenceRadioButton_Checked;
        rbWindowLaunchAction.Checked += WindowReferenceRadioButton_Checked;

        UpdateUIVisibility();
        PopulateLaunchActions();
    }

    public MouseActionDialog(MouseAction existingAction, string? currentScriptPath = null, ScriptEngine? scriptEngine = null)
        : this(currentScriptPath, scriptEngine)
    {
        // 既存のアクションの値を設定（編集モード）
        cmbClickType.SelectedIndex = existingAction.ClickType switch
        {
            MouseClickType.LeftClick => 0,
            MouseClickType.RightClick => 1,
            MouseClickType.DoubleClick => 2,
            MouseClickType.MiddleClick => 3,
            _ => 0
        };

        // クリック方法に応じて設定
        if (existingAction.Method == MouseAction.ClickMethod.Coordinate)
        {
            rbCoordinate.IsChecked = true;
            txtX.Text = existingAction.X.ToString();
            txtY.Text = existingAction.Y.ToString();
        }
        else
        {
            rbImageMatch.IsChecked = true;
            txtTemplatePath.Text = existingAction.TemplateImagePath ?? "";
            txtMatchThreshold.Text = existingAction.MatchThreshold.ToString("F2");
            txtTimeout.Text = existingAction.SearchTimeoutMs.ToString();
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
        }

        UpdateUIVisibility();
    }

    private void RadioButton_Checked(object sender, RoutedEventArgs e)
    {
        UpdateUIVisibility();
    }

    private void WindowReferenceRadioButton_Checked(object sender, RoutedEventArgs e)
    {
        UpdateWindowReferenceVisibility();
    }

    private void UpdateUIVisibility()
    {
        if (grpCoordinate == null || grpImageMatch == null) return;

        grpCoordinate.Visibility = rbCoordinate.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        grpImageMatch.Visibility = rbImageMatch.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
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
        System.Diagnostics.Debug.WriteLine($"Opening ScreenCaptureDialog with script path: {_currentScriptPath}");
        var captureDialog = new ScreenCaptureDialog(_currentScriptPath);
        captureDialog.ShowDialog();

        System.Diagnostics.Debug.WriteLine($"ScreenCaptureDialog closed. SavedImagePath = {captureDialog.SavedImagePath}");

        if (!string.IsNullOrEmpty(captureDialog.SavedImagePath))
        {
            txtTemplatePath.Text = captureDialog.SavedImagePath;
            System.Diagnostics.Debug.WriteLine($"Path set to textbox: {txtTemplatePath.Text}");
            MessageBox.Show(
                $"画像を保存しました:\n{captureDialog.SavedImagePath}",
                "保存完了",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );
        }
        else
        {
            System.Diagnostics.Debug.WriteLine("SavedImagePath is null or empty - no path set");
        }
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        var clickType = cmbClickType.SelectedIndex switch
        {
            0 => MouseClickType.LeftClick,
            1 => MouseClickType.RightClick,
            2 => MouseClickType.DoubleClick,
            3 => MouseClickType.MiddleClick,
            _ => MouseClickType.LeftClick
        };

        if (rbCoordinate.IsChecked == true)
        {
            // 座標指定モード
            if (!int.TryParse(txtX.Text, out int x) || x < 0)
            {
                MessageBox.Show("X座標が無効です", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!int.TryParse(txtY.Text, out int y) || y < 0)
            {
                MessageBox.Show("Y座標が無効です", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            Action = new MouseAction
            {
                Method = MouseAction.ClickMethod.Coordinate,
                X = x,
                Y = y,
                ClickType = clickType
            };
        }
        else
        {
            // 画像認識モード
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

            if (!double.TryParse(txtMatchThreshold.Text, out double threshold) || threshold < 0 || threshold > 1)
            {
                MessageBox.Show("一致率は0.0～1.0の範囲で指定してください", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!int.TryParse(txtTimeout.Text, out int timeout) || timeout < 0)
            {
                MessageBox.Show("タイムアウト時間が無効です", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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

            Action = new MouseAction
            {
                Method = MouseAction.ClickMethod.ImageMatch,
                TemplateImagePath = txtTemplatePath.Text,
                MatchThreshold = threshold,
                SearchTimeoutMs = timeout,
                UseMultiScale = chkMultiScale.IsChecked == true,
                WindowReference = windowReference,
                TargetWindowTitle = targetWindowTitle,
                TargetLaunchActionIndex = targetLaunchActionIndex,
                SearchAreaX = searchX,
                SearchAreaY = searchY,
                SearchAreaWidth = searchWidth,
                SearchAreaHeight = searchHeight,
                ClickType = clickType
            };
        }

        DialogResult = true;
        Close();
    }

    private async void BtnGetMousePos_Click(object sender, RoutedEventArgs e)
    {
        btnGetMousePos.IsEnabled = false;
        btnGetMousePos.Content = "3秒後にマウス位置を取得...";

        await Task.Delay(3000);

        GetCursorPos(out POINT point);
        txtX.Text = point.X.ToString();
        txtY.Text = point.Y.ToString();

        btnGetMousePos.IsEnabled = true;
        btnGetMousePos.Content = "マウス位置を取得 (3秒後)";
    }
}
