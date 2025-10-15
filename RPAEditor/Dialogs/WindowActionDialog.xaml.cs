using System.Windows;
using Microsoft.Win32;
using RPACore;
using RPACore.Actions;

namespace RPAEditor;

public partial class WindowActionDialog : Window
{
    public WindowAction Action { get; private set; }
    private List<IAction> _allActions;
    private int _currentActionIndex; // 編集時の現在のアクション番号（1始まり）

    public WindowActionDialog(List<IAction> allActions, int currentActionIndex = -1)
    {
        InitializeComponent();
        Action = new WindowAction();
        _allActions = allActions;
        _currentActionIndex = currentActionIndex;

        cmbActionType.SelectionChanged += CmbActionType_SelectionChanged;
        rbByTitle.Checked += RbReferenceType_Changed;
        rbByLaunchAction.Checked += RbReferenceType_Changed;
        btnBrowse.Click += BtnBrowse_Click;
        btnOK.Click += BtnOK_Click;

        // 初期選択
        cmbActionType.SelectedIndex = 0;
        LoadLaunchActions();
    }

    public WindowActionDialog(List<IAction> allActions, WindowAction existingAction, int currentActionIndex)
        : this(allActions, currentActionIndex)
    {
        // 既存のアクションの値を設定（編集モード）
        SelectActionType(existingAction.ActionType);

        if (existingAction.ActionType == WindowActionType.Launch)
        {
            txtExecutablePath.Text = existingAction.ExecutablePath;
            txtArguments.Text = existingAction.Arguments;
            txtWaitAfterLaunch.Text = existingAction.WaitAfterLaunch.ToString();
        }
        else
        {
            // 参照方法を設定
            if (existingAction.ReferenceType == WindowReferenceType.LaunchActionIndex)
            {
                rbByLaunchAction.IsChecked = true;
                SelectLaunchActionByIndex(existingAction.LaunchActionIndex);
            }
            else
            {
                rbByTitle.IsChecked = true;
                txtWindowTitle.Text = existingAction.WindowTitle;
            }
        }
    }

    private void SelectActionType(WindowActionType actionType)
    {
        string tag = actionType.ToString();
        foreach (System.Windows.Controls.ComboBoxItem item in cmbActionType.Items)
        {
            if (item.Tag?.ToString() == tag)
            {
                cmbActionType.SelectedItem = item;
                break;
            }
        }
    }

    private void CmbActionType_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (cmbActionType.SelectedItem is not System.Windows.Controls.ComboBoxItem selectedItem)
            return;

        string tag = selectedItem.Tag?.ToString() ?? "";

        // UIの表示/非表示を切り替え
        if (tag == "Launch")
        {
            pnlLaunch.Visibility = Visibility.Visible;
            pnlWindowTitle.Visibility = Visibility.Collapsed;
            txtExample.Text = "例: C:\\Program Files\\Notepad\\notepad.exe を起動";
        }
        else
        {
            pnlLaunch.Visibility = Visibility.Collapsed;
            pnlWindowTitle.Visibility = Visibility.Visible;

            txtExample.Text = tag switch
            {
                "Activate" => "例: タイトルに「メモ帳」を含むウィンドウをアクティブにする",
                "Maximize" => "例: タイトルに「Excel」を含むウィンドウを最大化",
                "Minimize" => "例: タイトルに「Chrome」を含むウィンドウを最小化",
                "Restore" => "例: 最小化されたウィンドウを元のサイズに戻す",
                "Close" => "例: タイトルに「新しいタブ」を含むウィンドウを閉じる",
                _ => "操作の種類を選択してください"
            };
        }
    }

    private void LoadLaunchActions()
    {
        cmbLaunchAction.Items.Clear();

        // 現在のアクションより前の起動アクションのみ表示
        int limit = _currentActionIndex > 0 ? _currentActionIndex - 1 : _allActions.Count;

        for (int i = 0; i < limit && i < _allActions.Count; i++)
        {
            if (_allActions[i] is WindowAction wa && wa.ActionType == WindowActionType.Launch)
            {
                var item = new System.Windows.Controls.ComboBoxItem
                {
                    Content = $"#{i + 1}: {wa.Name}",
                    Tag = i + 1 // 1始まりの番号を保存
                };
                cmbLaunchAction.Items.Add(item);
            }
        }

        if (cmbLaunchAction.Items.Count > 0)
            cmbLaunchAction.SelectedIndex = 0;
    }

    private void SelectLaunchActionByIndex(int actionIndex)
    {
        foreach (System.Windows.Controls.ComboBoxItem item in cmbLaunchAction.Items)
        {
            if (item.Tag is int index && index == actionIndex)
            {
                cmbLaunchAction.SelectedItem = item;
                break;
            }
        }
    }

    private void RbReferenceType_Changed(object sender, RoutedEventArgs e)
    {
        if (rbByTitle.IsChecked == true)
        {
            pnlTitleInput.Visibility = Visibility.Visible;
            pnlLaunchActionInput.Visibility = Visibility.Collapsed;
        }
        else if (rbByLaunchAction.IsChecked == true)
        {
            pnlTitleInput.Visibility = Visibility.Collapsed;
            pnlLaunchActionInput.Visibility = Visibility.Visible;
        }
    }

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*",
            Title = "起動する実行ファイルを選択"
        };

        if (dialog.ShowDialog() == true)
        {
            txtExecutablePath.Text = dialog.FileName;
        }
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        if (cmbActionType.SelectedItem is not System.Windows.Controls.ComboBoxItem selectedItem)
        {
            MessageBox.Show("操作の種類を選択してください", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        string tag = selectedItem.Tag?.ToString() ?? "";
        var actionType = Enum.Parse<WindowActionType>(tag);

        if (actionType == WindowActionType.Launch)
        {
            // アプリ起動の検証
            if (string.IsNullOrWhiteSpace(txtExecutablePath.Text))
            {
                MessageBox.Show("実行ファイルを指定してください", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!int.TryParse(txtWaitAfterLaunch.Text, out int waitTime) || waitTime < 0)
            {
                MessageBox.Show("待機時間は0以上の整数で指定してください", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            Action = new WindowAction(WindowActionType.Launch, txtExecutablePath.Text)
            {
                Arguments = txtArguments.Text,
                WaitAfterLaunch = waitTime
            };
        }
        else
        {
            // ウィンドウ操作の検証と作成
            if (rbByLaunchAction.IsChecked == true)
            {
                // 起動アクションで指定
                if (cmbLaunchAction.SelectedItem is not System.Windows.Controls.ComboBoxItem launchItem)
                {
                    MessageBox.Show("起動アクションを選択してください", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                int launchActionIndex = (int)launchItem.Tag!;
                Action = new WindowAction
                {
                    ActionType = actionType,
                    ReferenceType = WindowReferenceType.LaunchActionIndex,
                    LaunchActionIndex = launchActionIndex
                };
            }
            else
            {
                // ウィンドウタイトルで指定
                if (string.IsNullOrWhiteSpace(txtWindowTitle.Text))
                {
                    MessageBox.Show("ウィンドウタイトルを入力してください", "エラー",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Action = new WindowAction
                {
                    ActionType = actionType,
                    ReferenceType = WindowReferenceType.WindowTitle,
                    WindowTitle = txtWindowTitle.Text
                };
            }
        }

        DialogResult = true;
        Close();
    }
}
