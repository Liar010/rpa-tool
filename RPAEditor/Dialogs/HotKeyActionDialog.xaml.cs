using System.Windows;
using System.Windows.Controls;
using RPACore;
using RPACore.Actions;

namespace RPAEditor;

public partial class HotKeyActionDialog : Window
{
    public KeyboardAction Action { get; private set; }

    public HotKeyActionDialog()
    {
        InitializeComponent();
        Action = new KeyboardAction();

        // ラジオボタンのイベント登録
        rbAlphaNum.Checked += OnKeyTypeChanged;
        rbSymbol.Checked += OnKeyTypeChanged;
        rbFunction.Checked += OnKeyTypeChanged;
        rbSpecial.Checked += OnKeyTypeChanged;

        // 修飾キーとキー選択の変更イベント
        chkCtrl.Checked += OnKeySelectionChanged;
        chkCtrl.Unchecked += OnKeySelectionChanged;
        chkShift.Checked += OnKeySelectionChanged;
        chkShift.Unchecked += OnKeySelectionChanged;
        chkAlt.Checked += OnKeySelectionChanged;
        chkAlt.Unchecked += OnKeySelectionChanged;
        chkWin.Checked += OnKeySelectionChanged;
        chkWin.Unchecked += OnKeySelectionChanged;
        cmbKey.SelectionChanged += OnKeySelectionChanged;

        // 初期表示を設定
        LoadAlphaNumKeys();

        btnOK.Click += BtnOK_Click;
    }

    public HotKeyActionDialog(KeyboardAction existingAction) : this()
    {
        // 既存のアクションの値を設定（編集モード）
        chkCtrl.IsChecked = existingAction.ModifierKeys.Contains(VirtualKey.Control);
        chkShift.IsChecked = existingAction.ModifierKeys.Contains(VirtualKey.Shift);
        chkAlt.IsChecked = existingAction.ModifierKeys.Contains(VirtualKey.Alt);
        chkWin.IsChecked = existingAction.ModifierKeys.Contains(VirtualKey.LWin);

        // キーの設定 - 適切なラジオボタンを選択してからキーを設定
        SelectAppropriateKeyType(existingAction.Key);
        SelectKeyInComboBox(existingAction.Key.ToString());
    }

    private void OnKeyTypeChanged(object sender, RoutedEventArgs e)
    {
        if (rbAlphaNum.IsChecked == true)
            LoadAlphaNumKeys();
        else if (rbSymbol.IsChecked == true)
            LoadSymbolKeys();
        else if (rbFunction.IsChecked == true)
            LoadFunctionKeys();
        else if (rbSpecial.IsChecked == true)
            LoadSpecialKeys();
    }

    private void OnKeySelectionChanged(object sender, RoutedEventArgs e)
    {
        CheckSecurityRestriction();
    }

    private void OnKeySelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        CheckSecurityRestriction();
    }

    private void CheckSecurityRestriction()
    {
        // Ctrl+Alt+Delete の組み合わせをチェック
        bool isCtrl = chkCtrl.IsChecked == true;
        bool isAlt = chkAlt.IsChecked == true;
        bool isDelete = (cmbKey.SelectedItem as ComboBoxItem)?.Tag?.ToString() == "Delete";

        if (isCtrl && isAlt && isDelete)
        {
            warningBorder.Visibility = Visibility.Visible;
            txtWarning.Text = "Ctrl+Alt+DeleteはWindowsのセキュリティ機能により送信できません。\nこの組み合わせは実行しても動作しません。";
        }
        else
        {
            warningBorder.Visibility = Visibility.Collapsed;
        }
    }

    private void LoadAlphaNumKeys()
    {
        cmbKey.Items.Clear();

        // A-Z
        for (char c = 'A'; c <= 'Z'; c++)
        {
            cmbKey.Items.Add(new ComboBoxItem { Content = c.ToString(), Tag = c.ToString() });
        }

        cmbKey.Items.Add(new Separator());

        // 0-9
        for (int i = 0; i <= 9; i++)
        {
            cmbKey.Items.Add(new ComboBoxItem { Content = i.ToString(), Tag = $"D{i}" });
        }

        if (cmbKey.Items.Count > 0)
            cmbKey.SelectedIndex = 0;
    }

    private void LoadSymbolKeys()
    {
        cmbKey.Items.Clear();

        cmbKey.Items.Add(new ComboBoxItem { Content = "; (セミコロン - Ctrl+; で日付)", Tag = "OemSemicolon" });
        cmbKey.Items.Add(new ComboBoxItem { Content = ": (コロン - Ctrl+: で時刻)", Tag = "OemSemicolon" }); // Shift+;
        cmbKey.Items.Add(new ComboBoxItem { Content = ", (カンマ)", Tag = "OemComma" });
        cmbKey.Items.Add(new ComboBoxItem { Content = ". (ピリオド)", Tag = "OemPeriod" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "/ (スラッシュ)", Tag = "OemQuestion" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "- (マイナス)", Tag = "OemMinus" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "= (イコール)", Tag = "OemPlus" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "[ (左角かっこ)", Tag = "OemOpenBrackets" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "] (右角かっこ)", Tag = "OemCloseBrackets" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "\\ (バックスラッシュ)", Tag = "OemPipe" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "' (シングルクォート)", Tag = "OemQuotes" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "` (バッククォート)", Tag = "OemTilde" });

        if (cmbKey.Items.Count > 0)
            cmbKey.SelectedIndex = 0;
    }

    private void LoadFunctionKeys()
    {
        cmbKey.Items.Clear();

        for (int i = 1; i <= 12; i++)
        {
            cmbKey.Items.Add(new ComboBoxItem { Content = $"F{i}", Tag = $"F{i}" });
        }

        if (cmbKey.Items.Count > 0)
            cmbKey.SelectedIndex = 0;
    }

    private void LoadSpecialKeys()
    {
        cmbKey.Items.Clear();

        cmbKey.Items.Add(new ComboBoxItem { Content = "Enter", Tag = "Enter" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "Escape", Tag = "Escape" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "Tab", Tag = "Tab" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "Space (スペース)", Tag = "Space" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "Backspace", Tag = "Back" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "Delete", Tag = "Delete" });
        cmbKey.Items.Add(new Separator());
        cmbKey.Items.Add(new ComboBoxItem { Content = "↑ (上矢印)", Tag = "Up" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "↓ (下矢印)", Tag = "Down" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "← (左矢印)", Tag = "Left" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "→ (右矢印)", Tag = "Right" });
        cmbKey.Items.Add(new Separator());
        cmbKey.Items.Add(new ComboBoxItem { Content = "Home", Tag = "Home" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "End", Tag = "End" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "PageUp", Tag = "PageUp" });
        cmbKey.Items.Add(new ComboBoxItem { Content = "PageDown", Tag = "PageDown" });

        if (cmbKey.Items.Count > 0)
            cmbKey.SelectedIndex = 0;
    }

    private void SelectAppropriateKeyType(VirtualKey key)
    {
        // キーの種類に応じて適切なラジオボタンを選択
        if (key >= VirtualKey.A && key <= VirtualKey.Z || key >= VirtualKey.D0 && key <= VirtualKey.D9)
        {
            rbAlphaNum.IsChecked = true;
        }
        else if (key >= VirtualKey.F1 && key <= VirtualKey.F12)
        {
            rbFunction.IsChecked = true;
        }
        else if (key == VirtualKey.OemSemicolon || key == VirtualKey.OemPlus ||
                 key == VirtualKey.OemComma || key == VirtualKey.OemMinus ||
                 key == VirtualKey.OemPeriod || key == VirtualKey.OemQuestion ||
                 key == VirtualKey.OemTilde || key == VirtualKey.OemOpenBrackets ||
                 key == VirtualKey.OemPipe || key == VirtualKey.OemCloseBrackets ||
                 key == VirtualKey.OemQuotes)
        {
            rbSymbol.IsChecked = true;
        }
        else
        {
            rbSpecial.IsChecked = true;
        }
    }

    private void SelectKeyInComboBox(string keyTag)
    {
        foreach (var item in cmbKey.Items)
        {
            if (item is ComboBoxItem comboItem && comboItem.Tag?.ToString() == keyTag)
            {
                cmbKey.SelectedItem = comboItem;
                break;
            }
        }
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        if (chkCtrl.IsChecked != true && chkShift.IsChecked != true &&
            chkAlt.IsChecked != true && chkWin.IsChecked != true)
        {
            MessageBox.Show("少なくとも1つの修飾キーを選択してください", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var modifiers = new List<VirtualKey>();
        if (chkCtrl.IsChecked == true) modifiers.Add(VirtualKey.Control);
        if (chkShift.IsChecked == true) modifiers.Add(VirtualKey.Shift);
        if (chkAlt.IsChecked == true) modifiers.Add(VirtualKey.Alt);
        if (chkWin.IsChecked == true) modifiers.Add(VirtualKey.LWin);

        var selectedKey = (cmbKey.SelectedItem as ComboBoxItem)?.Tag?.ToString();
        if (string.IsNullOrEmpty(selectedKey))
        {
            MessageBox.Show("キーを選択してください", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var key = ConvertTagToVirtualKey(selectedKey);

        Action = new KeyboardAction(key, modifiers.ToArray());
        DialogResult = true;
        Close();
    }

    private VirtualKey ConvertTagToVirtualKey(string tag)
    {
        return tag switch
        {
            // アルファベット
            "A" => VirtualKey.A, "B" => VirtualKey.B, "C" => VirtualKey.C, "D" => VirtualKey.D,
            "E" => VirtualKey.E, "F" => VirtualKey.F, "G" => VirtualKey.G, "H" => VirtualKey.H,
            "I" => VirtualKey.I, "J" => VirtualKey.J, "K" => VirtualKey.K, "L" => VirtualKey.L,
            "M" => VirtualKey.M, "N" => VirtualKey.N, "O" => VirtualKey.O, "P" => VirtualKey.P,
            "Q" => VirtualKey.Q, "R" => VirtualKey.R, "S" => VirtualKey.S, "T" => VirtualKey.T,
            "U" => VirtualKey.U, "V" => VirtualKey.V, "W" => VirtualKey.W, "X" => VirtualKey.X,
            "Y" => VirtualKey.Y, "Z" => VirtualKey.Z,
            // 数字
            "D0" => VirtualKey.D0, "D1" => VirtualKey.D1, "D2" => VirtualKey.D2, "D3" => VirtualKey.D3,
            "D4" => VirtualKey.D4, "D5" => VirtualKey.D5, "D6" => VirtualKey.D6, "D7" => VirtualKey.D7,
            "D8" => VirtualKey.D8, "D9" => VirtualKey.D9,
            // 記号
            "OemSemicolon" => VirtualKey.OemSemicolon,
            "OemPlus" => VirtualKey.OemPlus,
            "OemComma" => VirtualKey.OemComma,
            "OemMinus" => VirtualKey.OemMinus,
            "OemPeriod" => VirtualKey.OemPeriod,
            "OemQuestion" => VirtualKey.OemQuestion,
            "OemTilde" => VirtualKey.OemTilde,
            "OemOpenBrackets" => VirtualKey.OemOpenBrackets,
            "OemPipe" => VirtualKey.OemPipe,
            "OemCloseBrackets" => VirtualKey.OemCloseBrackets,
            "OemQuotes" => VirtualKey.OemQuotes,
            // ファンクション
            "F1" => VirtualKey.F1, "F2" => VirtualKey.F2, "F3" => VirtualKey.F3, "F4" => VirtualKey.F4,
            "F5" => VirtualKey.F5, "F6" => VirtualKey.F6, "F7" => VirtualKey.F7, "F8" => VirtualKey.F8,
            "F9" => VirtualKey.F9, "F10" => VirtualKey.F10, "F11" => VirtualKey.F11, "F12" => VirtualKey.F12,
            // 特殊キー
            "Enter" => VirtualKey.Enter, "Escape" => VirtualKey.Escape, "Tab" => VirtualKey.Tab,
            "Space" => VirtualKey.Space, "Back" => VirtualKey.Back, "Delete" => VirtualKey.Delete,
            "Up" => VirtualKey.Up, "Down" => VirtualKey.Down, "Left" => VirtualKey.Left, "Right" => VirtualKey.Right,
            "Home" => VirtualKey.Home, "End" => VirtualKey.End, "PageUp" => VirtualKey.PageUp, "PageDown" => VirtualKey.PageDown,
            _ => VirtualKey.A
        };
    }
}
