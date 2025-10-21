using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using RPACore;
using RPACore.Actions;
using RPAEditor.Views;

namespace RPAEditor.Dialogs;

public partial class LoopEndActionDialog : Window
{
    public LoopEndAction Action { get; private set; }
    private readonly ScriptEngine? _scriptEngine;

    public LoopEndActionDialog(ScriptEngine? scriptEngine = null)
    {
        InitializeComponent();
        Action = new LoopEndAction();
        _scriptEngine = scriptEngine;

        btnOK.Click += BtnOK_Click;
        btnCancel.Click += (s, e) => DialogResult = false;

        cmbEndConditionType.SelectionChanged += CmbEndConditionType_SelectionChanged;

        LoadLoopStartActions();
    }

    public LoopEndActionDialog(LoopEndAction existingAction, ScriptEngine? scriptEngine = null) : this(scriptEngine)
    {
        txtLeftValue.Text = existingAction.LeftValue;
        txtRightValue.Text = existingAction.RightValue;
        txtMaxIterations.Text = existingAction.MaxIterations.ToString();

        // Set end condition type
        cmbEndConditionType.SelectedIndex = existingAction.EndConditionType switch
        {
            LoopEndConditionType.Always => 0,
            LoopEndConditionType.IfEqual => 1,
            LoopEndConditionType.IfNotEqual => 2,
            LoopEndConditionType.IfGreaterThan => 3,
            LoopEndConditionType.IfLessThan => 4,
            LoopEndConditionType.IfEmpty => 5,
            LoopEndConditionType.IfNotEmpty => 6,
            LoopEndConditionType.MaxIterations => 7,
            _ => 0
        };

        // Select loop start action
        var item = cmbLoopStartAction.Items.Cast<ActionListItem>()
            .FirstOrDefault(x => x.Index == existingAction.LoopStartActionIndex);
        if (item != null)
        {
            cmbLoopStartAction.SelectedItem = item;
        }
    }

    private void LoadLoopStartActions()
    {
        if (_scriptEngine == null)
            return;

        var loopStartActions = new List<ActionListItem>();

        for (int i = 0; i < _scriptEngine.Actions.Count; i++)
        {
            if (_scriptEngine.Actions[i] is LoopStartAction loopStart)
            {
                loopStartActions.Add(new ActionListItem
                {
                    Index = i + 1,
                    Action = loopStart
                });
            }
        }

        cmbLoopStartAction.ItemsSource = loopStartActions;
        if (loopStartActions.Count > 0)
        {
            cmbLoopStartAction.SelectedIndex = 0;
        }
    }

    private void CmbEndConditionType_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (pnlLeftValue == null || pnlRightValue == null || pnlMaxIterations == null)
            return;

        var selectedItem = (ComboBoxItem)cmbEndConditionType.SelectedItem;
        var conditionTag = selectedItem.Tag.ToString();

        switch (conditionTag)
        {
            case "Always":
                // 常に終了の場合は何も表示しない
                pnlLeftValue.Visibility = Visibility.Collapsed;
                pnlRightValue.Visibility = Visibility.Collapsed;
                pnlMaxIterations.Visibility = Visibility.Collapsed;
                break;

            case "IfEqual":
            case "IfNotEqual":
            case "IfGreaterThan":
            case "IfLessThan":
                // 比較条件の場合は左辺値と右辺値を表示
                pnlLeftValue.Visibility = Visibility.Visible;
                pnlRightValue.Visibility = Visibility.Visible;
                pnlMaxIterations.Visibility = Visibility.Collapsed;
                break;

            case "IfEmpty":
            case "IfNotEmpty":
                // 空チェックの場合は左辺値のみ表示
                pnlLeftValue.Visibility = Visibility.Visible;
                pnlRightValue.Visibility = Visibility.Collapsed;
                pnlMaxIterations.Visibility = Visibility.Collapsed;
                break;

            case "MaxIterations":
                // 最大繰り返し回数の場合
                pnlLeftValue.Visibility = Visibility.Collapsed;
                pnlRightValue.Visibility = Visibility.Collapsed;
                pnlMaxIterations.Visibility = Visibility.Visible;
                break;
        }
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        // Validation
        if (cmbLoopStartAction.SelectedItem == null)
        {
            MessageBox.Show("対応するループ開始アクションを選択してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var selectedCondition = (ComboBoxItem)cmbEndConditionType.SelectedItem;
        var conditionTag = selectedCondition.Tag.ToString();

        // 条件に応じたバリデーション
        if (conditionTag == "IfEqual" || conditionTag == "IfNotEqual" ||
            conditionTag == "IfGreaterThan" || conditionTag == "IfLessThan")
        {
            if (string.IsNullOrWhiteSpace(txtLeftValue.Text) || string.IsNullOrWhiteSpace(txtRightValue.Text))
            {
                MessageBox.Show("左辺値と右辺値を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        else if (conditionTag == "IfEmpty" || conditionTag == "IfNotEmpty")
        {
            if (string.IsNullOrWhiteSpace(txtLeftValue.Text))
            {
                MessageBox.Show("左辺値を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        else if (conditionTag == "MaxIterations")
        {
            if (!int.TryParse(txtMaxIterations.Text, out int maxIter) || maxIter <= 0)
            {
                MessageBox.Show("有効な最大繰り返し回数を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        // Get values
        var loopStartItem = (ActionListItem)cmbLoopStartAction.SelectedItem;
        int loopStartIndex = loopStartItem.Index; // Already 1-based

        var endConditionType = conditionTag switch
        {
            "IfEqual" => LoopEndConditionType.IfEqual,
            "IfNotEqual" => LoopEndConditionType.IfNotEqual,
            "IfGreaterThan" => LoopEndConditionType.IfGreaterThan,
            "IfLessThan" => LoopEndConditionType.IfLessThan,
            "IfEmpty" => LoopEndConditionType.IfEmpty,
            "IfNotEmpty" => LoopEndConditionType.IfNotEmpty,
            "MaxIterations" => LoopEndConditionType.MaxIterations,
            _ => LoopEndConditionType.Always
        };

        int maxIterations = 100;
        if (conditionTag == "MaxIterations")
        {
            int.TryParse(txtMaxIterations.Text, out maxIterations);
        }

        Action = new LoopEndAction
        {
            LoopStartActionIndex = loopStartIndex,
            EndConditionType = endConditionType,
            LeftValue = txtLeftValue.Text.Trim(),
            RightValue = txtRightValue.Text.Trim(),
            MaxIterations = maxIterations
        };

        DialogResult = true;
        Close();
    }
}
