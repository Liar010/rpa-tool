using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using RPAEditor.Services;

namespace RPAEditor.Views;

public partial class SettingsPage : UserControl
{
    public event EventHandler? BackRequested;

    private readonly ObservableCollection<CustomVariableItem> _customVariables;

    public SettingsPage()
    {
        InitializeComponent();

        _customVariables = new ObservableCollection<CustomVariableItem>();

        btnBack.Click += BtnBack_Click;
        btnAddVariable.Click += BtnAddVariable_Click;
        btnEditVariable.Click += BtnEditVariable_Click;
        btnDeleteVariable.Click += BtnDeleteVariable_Click;
        txtPreviewInput.TextChanged += TxtPreviewInput_TextChanged;

        LoadBuiltInVariables();
        LoadCustomVariables();
    }

    private void BtnBack_Click(object sender, RoutedEventArgs e)
    {
        BackRequested?.Invoke(this, EventArgs.Empty);
    }

    private void LoadBuiltInVariables()
    {
        var variables = FileNameTemplateHelper.GetBuiltInVariables();
        dgBuiltInVariables.ItemsSource = variables;
    }

    private void LoadCustomVariables()
    {
        _customVariables.Clear();
        var customVars = SettingsManager.Instance.CustomVariables;

        foreach (var kvp in customVars)
        {
            _customVariables.Add(new CustomVariableItem { Key = kvp.Key, Value = kvp.Value });
        }

        dgCustomVariables.ItemsSource = _customVariables;

        if (_customVariables.Any())
        {
            dgCustomVariables.Visibility = Visibility.Visible;
            txtNoCustomVariables.Visibility = Visibility.Collapsed;
        }
        else
        {
            dgCustomVariables.Visibility = Visibility.Collapsed;
            txtNoCustomVariables.Visibility = Visibility.Visible;
        }
    }

    private void BtnAddVariable_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Dialogs.CustomVariableDialog
        {
            Owner = Window.GetWindow(this)
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                SettingsManager.Instance.SetCustomVariable(dialog.VariableName, dialog.VariableValue);
                LoadCustomVariables();
                MessageBox.Show($"変数 '%{dialog.VariableName}%' を追加しました",
                    "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"変数の追加に失敗しました: {ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void BtnEditVariable_Click(object sender, RoutedEventArgs e)
    {
        if (dgCustomVariables.SelectedItem is CustomVariableItem selectedItem)
        {
            var dialog = new Dialogs.CustomVariableDialog(selectedItem.Key, selectedItem.Value)
            {
                Owner = Window.GetWindow(this)
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    // 変数名が変更された場合は古い変数を削除
                    if (selectedItem.Key != dialog.VariableName)
                    {
                        SettingsManager.Instance.RemoveCustomVariable(selectedItem.Key);
                    }

                    SettingsManager.Instance.SetCustomVariable(dialog.VariableName, dialog.VariableValue);
                    LoadCustomVariables();
                    MessageBox.Show($"変数 '%{dialog.VariableName}%' を更新しました",
                        "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"変数の更新に失敗しました: {ex.Message}",
                        "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        else
        {
            MessageBox.Show("編集する変数を選択してください", "確認",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void BtnDeleteVariable_Click(object sender, RoutedEventArgs e)
    {
        if (dgCustomVariables.SelectedItem is CustomVariableItem selectedItem)
        {
            var result = MessageBox.Show(
                $"変数 '%{selectedItem.Key}%' を削除しますか？",
                "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    SettingsManager.Instance.RemoveCustomVariable(selectedItem.Key);
                    LoadCustomVariables();
                    MessageBox.Show($"変数 '%{selectedItem.Key}%' を削除しました",
                        "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"変数の削除に失敗しました: {ex.Message}",
                        "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        else
        {
            MessageBox.Show("削除する変数を選択してください", "確認",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void TxtPreviewInput_TextChanged(object sender, TextChangedEventArgs e)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(txtPreviewInput.Text))
            {
                txtPreviewOutput.Text = "（入力してください）";
            }
            else
            {
                txtPreviewOutput.Text = FileNameTemplateHelper.GetPreview(txtPreviewInput.Text);
            }
        }
        catch (Exception ex)
        {
            txtPreviewOutput.Text = $"エラー: {ex.Message}";
        }
    }
}

public class CustomVariableItem
{
    public string Key { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}
