using System.Windows;
using RPACore;
using RPACore.Actions;

namespace RPAEditor;

public partial class KeyPressDialog : Window
{
    public KeyboardAction Action { get; private set; }

    public KeyPressDialog()
    {
        InitializeComponent();
        Action = new KeyboardAction();
        btnOK.Click += BtnOK_Click;
    }

    public KeyPressDialog(KeyboardAction existingAction) : this()
    {
        // 既存のアクションの値を設定（編集モード）
        var keyTag = existingAction.Key.ToString();

        // ComboBoxから該当するアイテムを探して選択
        foreach (System.Windows.Controls.ComboBoxItem item in cmbKey.Items)
        {
            if (item.Tag?.ToString() == keyTag)
            {
                cmbKey.SelectedItem = item;
                break;
            }
        }
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        var selectedKey = (cmbKey.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Tag?.ToString() ?? "Enter";

        var key = selectedKey switch
        {
            "Enter" => VirtualKey.Enter,
            "Escape" => VirtualKey.Escape,
            "Tab" => VirtualKey.Tab,
            "Space" => VirtualKey.Space,
            "Back" => VirtualKey.Back,
            "Delete" => VirtualKey.Delete,
            "Up" => VirtualKey.Up,
            "Down" => VirtualKey.Down,
            "Left" => VirtualKey.Left,
            "Right" => VirtualKey.Right,
            "Home" => VirtualKey.Home,
            "End" => VirtualKey.End,
            "PageUp" => VirtualKey.PageUp,
            "PageDown" => VirtualKey.PageDown,
            "F1" => VirtualKey.F1,
            "F2" => VirtualKey.F2,
            "F3" => VirtualKey.F3,
            "F4" => VirtualKey.F4,
            "F5" => VirtualKey.F5,
            "F6" => VirtualKey.F6,
            "F7" => VirtualKey.F7,
            "F8" => VirtualKey.F8,
            "F9" => VirtualKey.F9,
            "F10" => VirtualKey.F10,
            "F11" => VirtualKey.F11,
            "F12" => VirtualKey.F12,
            _ => VirtualKey.Enter
        };

        Action = new KeyboardAction(key);
        Action.ActionType = KeyboardActionType.PressKey; // 単発キー入力モード
        DialogResult = true;
        Close();
    }
}
