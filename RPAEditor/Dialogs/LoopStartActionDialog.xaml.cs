using System.Windows;
using RPACore.Actions;

namespace RPAEditor.Dialogs;

public partial class LoopStartActionDialog : Window
{
    public LoopStartAction Action { get; private set; }

    public LoopStartActionDialog()
    {
        InitializeComponent();
        Action = new LoopStartAction();

        btnOK.Click += BtnOK_Click;
        btnCancel.Click += (s, e) => DialogResult = false;
    }

    public LoopStartActionDialog(LoopStartAction existingAction) : this()
    {
        txtComment.Text = existingAction.Comment;
    }

    private void BtnOK_Click(object sender, RoutedEventArgs e)
    {
        Action = new LoopStartAction
        {
            Comment = txtComment.Text.Trim()
        };

        DialogResult = true;
        Close();
    }
}
