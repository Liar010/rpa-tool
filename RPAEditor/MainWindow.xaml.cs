using System.Windows;
using RPACore;

namespace RPAEditor;

public partial class MainWindow : Window
{
    private readonly ScriptEngine _scriptEngine;
    private Views.HomePage? _homePage;
    private Views.EditorPage? _editorPage;

    public MainWindow()
    {
        InitializeComponent();
        _scriptEngine = new ScriptEngine();

        // 起動時はホームページを表示
        ShowHomePage();
    }

    private void ShowHomePage()
    {
        _homePage = new Views.HomePage();
        _homePage.NewScriptRequested += HomePage_NewScriptRequested;
        _homePage.ScriptFileSelected += HomePage_ScriptFileSelected;

        contentMain.Content = _homePage;
    }

    private void ShowEditorPage(string? scriptFilePath = null)
    {
        _editorPage = new Views.EditorPage(_scriptEngine, this);
        _editorPage.BackRequested += EditorPage_BackRequested;

        // スクリプトファイルが指定されている場合は読み込み
        if (!string.IsNullOrEmpty(scriptFilePath))
        {
            _editorPage.LoadScript(scriptFilePath);
        }

        contentMain.Content = _editorPage;
    }

    private void HomePage_NewScriptRequested(object? sender, System.EventArgs e)
    {
        // 新規スクリプト作成（空のエディタを開く）
        _scriptEngine.ClearActions();
        ShowEditorPage();
    }

    private void HomePage_ScriptFileSelected(object? sender, string scriptFilePath)
    {
        // 既存スクリプトを開く
        ShowEditorPage(scriptFilePath);
    }

    private void EditorPage_BackRequested(object? sender, System.EventArgs e)
    {
        // ホームページに戻る
        _homePage?.RefreshRecentScripts();
        ShowHomePage();
    }
}
