using System;
using System.IO;

namespace RPACore
{
    /// <summary>
    /// スクリプトと画像ファイルのパス管理
    ///
    /// 画像ファイル管理方針:
    /// - 保存場所: Scripts/[スクリプト名]_images/
    /// - ファイル名: 自動採番（mouse_click_001.png等）
    /// - パス保存: 相対パス（./[スクリプト名]_images/xxx.png）
    /// - すべてのファイルはexe配置フォルダ内で管理される
    /// </summary>
    public static class ScriptPathManager
    {
        /// <summary>
        /// アプリケーションのベースディレクトリ
        /// </summary>
        public static string BaseDirectory => AppDomain.CurrentDomain.BaseDirectory;

        /// <summary>
        /// スクリプト保存用の基本フォルダ
        /// </summary>
        public static string ScriptsDirectory => Path.Combine(BaseDirectory, "Scripts");

        /// <summary>
        /// スクリプトファイルパスから画像フォルダパスを取得
        /// 新しい方式: Scripts/スクリプト名/images/
        /// </summary>
        /// <param name="scriptFilePath">スクリプトファイルのパス</param>
        /// <returns>画像フォルダパス（例: "Scripts/勤怠入力/images/"）</returns>
        public static string GetImageFolderPath(string scriptFilePath)
        {
            if (string.IsNullOrEmpty(scriptFilePath))
                return Path.Combine(ScriptsDirectory, "未保存", "images");

            // スクリプトファイルのディレクトリ = スクリプトフォルダ
            string scriptDirectory = Path.GetDirectoryName(scriptFilePath) ?? ScriptsDirectory;

            return Path.Combine(scriptDirectory, "images");
        }

        /// <summary>
        /// 画像フォルダが存在しなければ作成
        /// </summary>
        /// <param name="scriptFilePath">スクリプトファイルのパス</param>
        public static void EnsureImageFolderExists(string scriptFilePath)
        {
            string imageFolder = GetImageFolderPath(scriptFilePath);
            if (!Directory.Exists(imageFolder))
            {
                Directory.CreateDirectory(imageFolder);
            }
        }

        /// <summary>
        /// 画像ファイルの保存パスを生成（自動採番）
        /// </summary>
        /// <param name="scriptFilePath">スクリプトファイルのパス</param>
        /// <param name="prefix">ファイル名のプレフィックス（デフォルト: "template"）</param>
        /// <returns>画像ファイルの絶対パス（例: "Scripts/勤怠入力_images/mouse_click_001.png"）</returns>
        public static string GenerateImageFilePath(string scriptFilePath, string prefix = "template")
        {
            EnsureImageFolderExists(scriptFilePath);
            string imageFolder = GetImageFolderPath(scriptFilePath);

            int index = 1;
            string filePath;
            do
            {
                filePath = Path.Combine(imageFolder, $"{prefix}_{index:D3}.png");
                index++;
            } while (File.Exists(filePath));

            return filePath;
        }

        /// <summary>
        /// 絶対パスを相対パスに変換（JSON保存用）
        /// </summary>
        /// <param name="absolutePath">絶対パス</param>
        /// <param name="basePath">基準となるパス（通常はスクリプトファイルのディレクトリ）</param>
        /// <returns>相対パス（例: "./勤怠入力_images/apply.png"）</returns>
        public static string? ToRelativePath(string? absolutePath, string basePath)
        {
            if (string.IsNullOrEmpty(absolutePath))
                return null;

            try
            {
                // basePath がディレクトリパスであることを保証
                if (!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()) &&
                    !basePath.EndsWith(Path.AltDirectorySeparatorChar.ToString()))
                {
                    basePath += Path.DirectorySeparatorChar;
                }

                Uri baseUri = new Uri(basePath);
                Uri absoluteUri = new Uri(absolutePath);
                Uri relativeUri = baseUri.MakeRelativeUri(absoluteUri);

                string relativePath = Uri.UnescapeDataString(relativeUri.ToString());
                // Unix スタイルのパスを Windows スタイルに変換
                relativePath = relativePath.Replace('/', Path.DirectorySeparatorChar);

                return "." + Path.DirectorySeparatorChar + relativePath;
            }
            catch
            {
                // 変換に失敗した場合は絶対パスをそのまま返す
                return absolutePath;
            }
        }

        /// <summary>
        /// 相対パスを絶対パスに変換（JSON読み込み用）
        /// </summary>
        /// <param name="relativePath">相対パス</param>
        /// <param name="basePath">基準となるパス（通常はスクリプトファイルのディレクトリ）</param>
        /// <returns>絶対パス</returns>
        public static string? ToAbsolutePath(string? relativePath, string basePath)
        {
            if (string.IsNullOrEmpty(relativePath))
                return null;

            // すでに絶対パスの場合はそのまま返す
            if (Path.IsPathRooted(relativePath))
                return relativePath;

            // "./" または ".\" を除去
            string cleanPath = relativePath.TrimStart('.', '/', '\\');

            return Path.GetFullPath(Path.Combine(basePath, cleanPath));
        }

        /// <summary>
        /// Scripts フォルダが存在しなければ作成
        /// </summary>
        public static void EnsureScriptsFolderExists()
        {
            if (!Directory.Exists(ScriptsDirectory))
            {
                Directory.CreateDirectory(ScriptsDirectory);
            }
        }
    }
}
