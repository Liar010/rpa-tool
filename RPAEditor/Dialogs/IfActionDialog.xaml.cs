using System.Windows;
using System.Windows.Controls;
using RPACore.Actions;

namespace RPAEditor.Dialogs;

public partial class IfActionDialog : Window
{
    public IfAction Action { get; private set; }

    public IfActionDialog()
    {
        InitializeComponent();
        Action = new IfAction();

        btnOK.Click += BtnOK_Click;
        btnCancel.Click += (s, e) => DialogResult = false;

        cmbConditionType.SelectionChanged += CmbConditionType_SelectionChanged;
        cmbThenAction.SelectionChanged += CmbThenAction_SelectionChanged;
    }

    public IfActionDialog(IfAction existingAction) : this()
    {
        txtLeftValue.Text = existingAction.LeftValue;
        txtRightValue.Text = existingAction.RightValue;

        // Set condition type
        cmbConditionType.SelectedIndex = existingAction.ConditionType switch
        {
            ConditionType.Equal => 0,
            ConditionType.NotEqual => 1,
            ConditionType.GreaterThan => 2,
            ConditionType.GreaterThanOrEqual => 3,
            ConditionType.LessThan => 4,
            ConditionType.LessThanOrEqual => 5,
            ConditionType.Contains => 6,
            ConditionType.NotContains => 7,
            ConditionType.IsEmpty => 8,
            ConditionType.IsNotEmpty => 9,
            _ => 0
        };

        // Set then action
        cmbThenAction.SelectedIndex = existingAction.ThenAction switch
        {
            IfThenAction.Continue => 0,
            IfThenAction.SkipNext => 1,
            IfThenAction.JumpToAction => 2,
            IfThenAction.ExitScript => 3,
            _ => 0
        };

        txtJumpToActionIndex.Text = existingAction.JumpToActionIndex.ToString();
    }

    private void CmbConditionType_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (lblRightValue == null || txtRightValue == null)
            return;

        var selectedItem = (ComboBoxItem)cmbConditionType.SelectedItem;
        var conditionTag = selectedItem.Tag.ToString();

        // IsEmpty と IsNotEmpty の場合は右辺値を無効化
        if (conditionTag == "IsEmpty" || conditionTag == "IsNotEmpty")
        {
            lblRightValue.IsEnabled = false;
            txtRightValue.IsEnabled = false;
            txtRightValue.Text = string.Empty;
        }
        else
        {
            lblRightValue.IsEnabled = true;
            txtRightValue.IsEnabled = true;
        }
    }

    private void CmbThenAction_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (pnlJumpTo == null)
            return;

        var selectedItem = (ComboBoxItem)cmbThenAction.SelectedItem;
        var actionTag = selectedItem.Tag.ToString();

        // JumpToAction の場合のみジャンプ先を表示
        pnlJumpTo.Visibility = actionTag == "JumpToAction" ? Visibility.Visible : Visibility.Collapsed;
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        // Validation
        if (string.IsNullOrWhiteSpace(txtLeftValue.Text))
        {
            MessageBox.Show("左辺値を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var selectedCondition = (ComboBoxItem)cmbConditionType.SelectedItem;
        var conditionTag = selectedCondition.Tag.ToString();

        // IsEmpty と IsNotEmpty 以外は右辺値が必要
        if (conditionTag != "IsEmpty" && conditionTag != "IsNotEmpty")
        {
            if (string.IsNullOrWhiteSpace(txtRightValue.Text))
            {
                MessageBox.Show("右辺値を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        var selectedThen = (ComboBoxItem)cmbThenAction.SelectedItem;
        var thenTag = selectedThen.Tag.ToString();

        // JumpToAction の場合はジャンプ先が必要
        int jumpToIndex = 0;
        if (thenTag == "JumpToAction")
        {
            if (!int.TryParse(txtJumpToActionIndex.Text, out jumpToIndex) || jumpToIndex <= 0)
            {
                MessageBox.Show("有効なジャンプ先アクション番号を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        // Create action
        var conditionType = conditionTag switch
        {
            "NotEqual" => ConditionType.NotEqual,
            "GreaterThan" => ConditionType.GreaterThan,
            "GreaterThanOrEqual" => ConditionType.GreaterThanOrEqual,
            "LessThan" => ConditionType.LessThan,
            "LessThanOrEqual" => ConditionType.LessThanOrEqual,
            "Contains" => ConditionType.Contains,
            "NotContains" => ConditionType.NotContains,
            "IsEmpty" => ConditionType.IsEmpty,
            "IsNotEmpty" => ConditionType.IsNotEmpty,
            _ => ConditionType.Equal
        };

        var thenAction = thenTag switch
        {
            "SkipNext" => IfThenAction.SkipNext,
            "JumpToAction" => IfThenAction.JumpToAction,
            "ExitScript" => IfThenAction.ExitScript,
            _ => IfThenAction.Continue
        };

        Action = new IfAction
        {
            LeftValue = txtLeftValue.Text.Trim(),
            RightValue = txtRightValue.Text.Trim(),
            ConditionType = conditionType,
            ThenAction = thenAction,
            JumpToActionIndex = jumpToIndex
        };

        DialogResult = true;
        Close();
    }
}
