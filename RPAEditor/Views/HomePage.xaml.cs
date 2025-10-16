using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using RPAEditor.Services;

namespace RPAEditor.Views;

public partial class HomePage : UserControl
{
    public event EventHandler? NewScriptRequested;
    public event EventHandler<string>? ScriptFileSelected;

    private readonly RecentFilesManager _recentFilesManager;

    public HomePage()
    {
        InitializeComponent();

        _recentFilesManager = new RecentFilesManager();

        btnNewScript.Click += BtnNewScript_Click;
        lstRecentScripts.MouseDoubleClick += LstRecentScripts_MouseDoubleClick;

        LoadRecentScripts();
    }

    private void BtnNewScript_Click(object sender, RoutedEventArgs e)
    {
        NewScriptRequested?.Invoke(this, EventArgs.Empty);
    }

    private void LstRecentScripts_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (lstRecentScripts.SelectedItem is RecentFileInfo selectedScript)
        {
            ScriptFileSelected?.Invoke(this, selectedScript.FilePath);
        }
    }

    private void LoadRecentScripts()
    {
        try
        {
            var recentFiles = _recentFilesManager.GetRecentFiles();

            if (recentFiles.Any())
            {
                lstRecentScripts.ItemsSource = recentFiles;
                lstRecentScripts.Visibility = Visibility.Visible;
                txtNoScripts.Visibility = Visibility.Collapsed;
            }
            else
            {
                lstRecentScripts.ItemsSource = null;
                lstRecentScripts.Visibility = Visibility.Collapsed;
                txtNoScripts.Visibility = Visibility.Visible;
            }
        }
        catch (Exception ex)
        {
            // エラーが発生しても空のリストを表示
            Console.WriteLine($"Failed to load recent scripts: {ex.Message}");
            lstRecentScripts.ItemsSource = null;
            lstRecentScripts.Visibility = Visibility.Collapsed;
            txtNoScripts.Visibility = Visibility.Visible;
        }
    }

    public void RefreshRecentScripts()
    {
        LoadRecentScripts();
    }
}
