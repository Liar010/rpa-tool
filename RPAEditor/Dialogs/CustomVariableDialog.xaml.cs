using System;
using System.Text.RegularExpressions;
using System.Windows;

namespace RPAEditor.Dialogs;

public partial class CustomVariableDialog : Window
{
    public string VariableName { get; private set; } = string.Empty;
    public string VariableValue { get; private set; } = string.Empty;

    // 新規追加用コンストラクタ
    public CustomVariableDialog()
    {
        InitializeComponent();
        btnOK.Click += BtnOK_Click;
        btnCancel.Click += BtnCancel_Click;
    }

    // 編集用コンストラクタ
    public CustomVariableDialog(string variableName, string variableValue) : this()
    {
        txtVariableName.Text = variableName;
        txtVariableValue.Text = variableValue;
        Title = "カスタム変数の編集";
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        // バリデーション
        var name = txtVariableName.Text.Trim();
        var value = txtVariableValue.Text;

        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show("変数名を入力してください", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            txtVariableName.Focus();
            return;
        }

        // 変数名のバリデーション（英数字とアンダースコアのみ）
        if (!Regex.IsMatch(name, @"^[a-zA-Z0-9_]+$"))
        {
            MessageBox.Show("変数名は英数字とアンダースコア(_)のみ使用できます", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            txtVariableName.Focus();
            return;
        }

        // 組み込み変数名との重複チェック
        var builtInNames = new[] { "date", "time", "datetime", "timestamp", "user", "computer",
            "year", "month", "day", "hour", "minute", "second" };

        if (Array.Exists(builtInNames, x => x.Equals(name, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show($"'{name}' は組み込み変数名と重複しています。別の名前を使用してください",
                "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            txtVariableName.Focus();
            return;
        }

        if (string.IsNullOrWhiteSpace(value))
        {
            MessageBox.Show("値を入力してください", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            txtVariableValue.Focus();
            return;
        }

        VariableName = name;
        VariableValue = value;
        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
