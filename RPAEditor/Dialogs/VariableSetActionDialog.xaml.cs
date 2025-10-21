using System.Windows;
using System.Windows.Controls;
using RPACore.Actions;

namespace RPAEditor.Dialogs;

public partial class VariableSetActionDialog : Window
{
    public VariableSetAction Action { get; private set; }

    public VariableSetActionDialog()
    {
        InitializeComponent();
        Action = new VariableSetAction();

        btnOK.Click += BtnOK_Click;
        btnCancel.Click += (s, e) => DialogResult = false;
    }

    public VariableSetActionDialog(VariableSetAction existingAction) : this()
    {
        txtVariableName.Text = existingAction.VariableName;
        txtValue.Text = existingAction.Value;

        // Set operation type
        switch (existingAction.OperationType)
        {
            case VariableSetOperationType.Set:
                cmbOperationType.SelectedIndex = 0;
                break;
            case VariableSetOperationType.Add:
                cmbOperationType.SelectedIndex = 1;
                break;
            case VariableSetOperationType.Subtract:
                cmbOperationType.SelectedIndex = 2;
                break;
        }
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        // Validation
        if (string.IsNullOrWhiteSpace(txtVariableName.Text))
        {
            MessageBox.Show("変数名を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(txtValue.Text))
        {
            MessageBox.Show("値を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Validate variable name (alphanumeric and underscore only)
        if (!System.Text.RegularExpressions.Regex.IsMatch(txtVariableName.Text, @"^[a-zA-Z0-9_]+$"))
        {
            MessageBox.Show("変数名は半角英数字とアンダースコアのみ使用できます。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Get operation type
        var selectedItem = (ComboBoxItem)cmbOperationType.SelectedItem;
        var operationType = selectedItem.Tag.ToString() switch
        {
            "Add" => VariableSetOperationType.Add,
            "Subtract" => VariableSetOperationType.Subtract,
            _ => VariableSetOperationType.Set
        };

        Action = new VariableSetAction
        {
            VariableName = txtVariableName.Text.Trim(),
            Value = txtValue.Text.Trim(),
            OperationType = operationType
        };

        DialogResult = true;
        Close();
    }
}
