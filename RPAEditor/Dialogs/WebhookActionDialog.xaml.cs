using System;
using System.Net.Http;
using System.Text;
using System.Windows;
using RPACore.Actions;

namespace RPAEditor.Dialogs;

public partial class WebhookActionDialog : Window
{
    public WebhookAction Action { get; private set; }

    // 新規作成用コンストラクタ
    public WebhookActionDialog()
    {
        InitializeComponent();
        Action = new WebhookAction();

        cmbServiceType.SelectedIndex = 0; // Discord
        cmbServiceType.SelectionChanged += CmbServiceType_SelectionChanged;
        btnTest.Click += BtnTest_Click;
        btnOK.Click += BtnOK_Click;
        btnCancel.Click += BtnCancel_Click;

        UpdateServiceInfo();
    }

    // 編集用コンストラクタ
    public WebhookActionDialog(WebhookAction existingAction) : this()
    {
        Action = existingAction;

        // 既存の値を設定
        txtWebhookUrl.Text = existingAction.WebhookUrl;
        txtMessage.Text = existingAction.Message;
        txtCustomPayload.Text = existingAction.CustomPayload;

        cmbServiceType.SelectedIndex = existingAction.ServiceType switch
        {
            WebhookServiceType.Discord => 0,
            WebhookServiceType.Slack => 1,
            WebhookServiceType.Teams => 2,
            WebhookServiceType.GoogleChat => 3,
            WebhookServiceType.Custom => 4,
            _ => 0
        };

        Title = "Webhook通知の編集";
    }

    private void CmbServiceType_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        UpdateServiceInfo();
    }

    private void UpdateServiceInfo()
    {
        var selectedItem = cmbServiceType.SelectedItem as System.Windows.Controls.ComboBoxItem;
        if (selectedItem == null) return;

        var tag = selectedItem.Tag as string;

        switch (tag)
        {
            case "Discord":
                txtUrlLabel.Text = "Discord Webhook URL:";
                pnlCustomPayload.Visibility = Visibility.Collapsed;
                pnlServiceInfo.Visibility = Visibility.Visible;
                txtServiceInfo.Text = "Discord Webhook URLの取得方法:\n" +
                    "1. サーバー設定 → 連携サービス → ウェブフック\n" +
                    "2. 新しいウェブフック → ウェブフックURLをコピー\n" +
                    "3. メッセージには変数（%date%, %time%等）が使用できます";
                break;

            case "Slack":
                txtUrlLabel.Text = "Slack Webhook URL:";
                pnlCustomPayload.Visibility = Visibility.Collapsed;
                pnlServiceInfo.Visibility = Visibility.Visible;
                txtServiceInfo.Text = "Slack Webhook URLの取得方法:\n" +
                    "1. https://api.slack.com/apps → Create New App\n" +
                    "2. Incoming Webhooks → Activate → Add New Webhook\n" +
                    "3. Webhook URLをコピー\n" +
                    "4. メッセージには変数（%date%, %time%等）が使用できます";
                break;

            case "Teams":
                txtUrlLabel.Text = "Teams Webhook URL:";
                pnlCustomPayload.Visibility = Visibility.Collapsed;
                pnlServiceInfo.Visibility = Visibility.Visible;
                txtServiceInfo.Text = "Teams Webhook URLの取得方法:\n" +
                    "1. チャネル → コネクタ → Incoming Webhook\n" +
                    "2. 構成 → 名前入力 → 作成\n" +
                    "3. URLをコピー\n" +
                    "4. メッセージには変数（%date%, %time%等）が使用できます";
                break;

            case "GoogleChat":
                txtUrlLabel.Text = "Google Chat Webhook URL:";
                pnlCustomPayload.Visibility = Visibility.Collapsed;
                pnlServiceInfo.Visibility = Visibility.Visible;
                txtServiceInfo.Text = "Google Chat Webhook URLの取得方法:\n" +
                    "1. Google Chatスペース → 統合 → Webhookを管理\n" +
                    "2. Webhookを作成 → 名前入力 → 保存\n" +
                    "3. URLをコピー\n" +
                    "4. メッセージには変数（%date%, %time%等）が使用できます";
                break;

            case "Custom":
                txtUrlLabel.Text = "Webhook URL:";
                pnlCustomPayload.Visibility = Visibility.Visible;
                pnlServiceInfo.Visibility = Visibility.Collapsed;
                txtCustomPayload.Text = "{\n  \"message\": \"RPA完了通知\",\n  \"timestamp\": \"%datetime%\",\n  \"status\": \"success\"\n}";
                break;
        }
    }

    private async void BtnTest_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtWebhookUrl.Text))
        {
            MessageBox.Show("Webhook URL/トークンを入力してください", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(txtMessage.Text))
        {
            MessageBox.Show("メッセージを入力してください", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        btnTest.IsEnabled = false;
        btnTest.Content = "送信中...";

        try
        {
            var testAction = CreateActionFromInput();
            testAction.Message = $"[テスト送信] {testAction.Message}";

            var success = await testAction.ExecuteAsync();

            if (success)
            {
                MessageBox.Show("テスト送信が成功しました！", "成功",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show($"テスト送信に失敗しました。\n\nエラー: {testAction.LastError}",
                    "失敗", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"テスト送信中にエラーが発生しました:\n{ex.Message}", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            btnTest.IsEnabled = true;
            btnTest.Content = "テスト送信";
        }
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        // バリデーション
        if (string.IsNullOrWhiteSpace(txtWebhookUrl.Text))
        {
            MessageBox.Show("Webhook URL/トークンを入力してください", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            txtWebhookUrl.Focus();
            return;
        }

        if (string.IsNullOrWhiteSpace(txtMessage.Text))
        {
            MessageBox.Show("メッセージを入力してください", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            txtMessage.Focus();
            return;
        }

        Action = CreateActionFromInput();
        DialogResult = true;
        Close();
    }

    private WebhookAction CreateActionFromInput()
    {
        var selectedItem = cmbServiceType.SelectedItem as System.Windows.Controls.ComboBoxItem;
        var tag = selectedItem?.Tag as string;

        var serviceType = tag switch
        {
            "Discord" => WebhookServiceType.Discord,
            "Slack" => WebhookServiceType.Slack,
            "Teams" => WebhookServiceType.Teams,
            "GoogleChat" => WebhookServiceType.GoogleChat,
            "Custom" => WebhookServiceType.Custom,
            _ => WebhookServiceType.Custom
        };

        return new WebhookAction
        {
            WebhookUrl = txtWebhookUrl.Text.Trim(),
            ServiceType = serviceType,
            Message = txtMessage.Text,
            CustomPayload = txtCustomPayload.Text,
            TimeoutMs = 10000
        };
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
