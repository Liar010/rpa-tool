using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace RPACore
{
    /// <summary>
    /// 画像マッチングサービス
    /// OpenCvSharpを使用したテンプレートマッチング機能を提供
    ///
    /// 主な機能:
    /// - 単一スケールマッチング
    /// - マルチスケールマッチング（DPI/解像度変更対応）
    /// - スクリーンキャプチャ
    /// </summary>
    public class ImageMatchingService : IDisposable
    {
        /// <summary>
        /// マルチスケールマッチング時に使用するスケールのリスト
        /// 0.5倍～2.0倍まで10段階
        /// </summary>
        private static readonly double[] DefaultScales =
        {
            0.5, 0.75, 0.8, 0.9, 1.0, 1.1, 1.2, 1.25, 1.5, 2.0
        };

        private bool _disposed = false;

        /// <summary>
        /// 画面全体をキャプチャしてMatオブジェクトとして返す
        /// </summary>
        public Mat CaptureScreen()
        {
            // 仮想画面全体のサイズを取得
            int screenLeft = NativeMethods.GetSystemMetrics(76);   // SM_XVIRTUALSCREEN
            int screenTop = NativeMethods.GetSystemMetrics(77);    // SM_YVIRTUALSCREEN
            int screenWidth = NativeMethods.GetSystemMetrics(78);  // SM_CXVIRTUALSCREEN
            int screenHeight = NativeMethods.GetSystemMetrics(79); // SM_CYVIRTUALSCREEN

            using var bitmap = new Bitmap(screenWidth, screenHeight, PixelFormat.Format24bppRgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(screenLeft, screenTop, 0, 0,
                    new System.Drawing.Size(screenWidth, screenHeight));
            }

            // BitmapをMatに変換
            return BitmapConverter.ToMat(bitmap);
        }

        /// <summary>
        /// 指定領域をキャプチャしてMatオブジェクトとして返す
        /// </summary>
        public Mat CaptureScreen(int x, int y, int width, int height)
        {
            using var bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
            }

            return BitmapConverter.ToMat(bitmap);
        }

        /// <summary>
        /// 画像ファイルを探してクリック位置を返す
        /// </summary>
        /// <param name="templatePath">テンプレート画像のパス</param>
        /// <param name="threshold">一致率の閾値（0.0～1.0）</param>
        /// <param name="useMultiScale">マルチスケールマッチングを使用するか</param>
        /// <param name="searchArea">検索範囲（nullの場合は全画面）</param>
        /// <returns>見つかった場合はクリック位置、見つからない場合はnull</returns>
        public System.Drawing.Point? FindImage(
            string templatePath,
            double threshold = 0.8,
            bool useMultiScale = true,
            Rectangle? searchArea = null)
        {
            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"テンプレート画像が見つかりません: {templatePath}");
            }

            using var template = Cv2.ImRead(templatePath, ImreadModes.Color);
            if (template.Empty())
            {
                throw new InvalidOperationException($"画像の読み込みに失敗しました: {templatePath}");
            }

            using var screenshot = searchArea.HasValue
                ? CaptureScreen(searchArea.Value.X, searchArea.Value.Y,
                    searchArea.Value.Width, searchArea.Value.Height)
                : CaptureScreen();

            if (useMultiScale)
            {
                var result = FindImageMultiScale(screenshot, template, threshold);
                if (result.HasValue && searchArea.HasValue)
                {
                    // 検索範囲のオフセットを加算
                    return new System.Drawing.Point(
                        result.Value.X + searchArea.Value.X,
                        result.Value.Y + searchArea.Value.Y
                    );
                }
                return result;
            }
            else
            {
                var result = FindImageSingleScale(screenshot, template, threshold);
                if (result.HasValue && searchArea.HasValue)
                {
                    return new System.Drawing.Point(
                        result.Value.X + searchArea.Value.X,
                        result.Value.Y + searchArea.Value.Y
                    );
                }
                return result;
            }
        }

        /// <summary>
        /// タイムアウト付きで画像を探す
        /// </summary>
        public async Task<System.Drawing.Point?> FindImageWithTimeoutAsync(
            string templatePath,
            double threshold = 0.8,
            int timeoutMs = 5000,
            bool useMultiScale = true,
            Rectangle? searchArea = null)
        {
            var startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalMilliseconds < timeoutMs)
            {
                var result = FindImage(templatePath, threshold, useMultiScale, searchArea);
                if (result.HasValue)
                {
                    return result;
                }

                // 少し待機してから再試行
                await Task.Delay(100);
            }

            return null;
        }

        /// <summary>
        /// マルチスケールマッチング
        /// 複数のスケールでテンプレート画像を探索し、最も一致率の高い位置を返す
        /// </summary>
        private System.Drawing.Point? FindImageMultiScale(Mat screenshot, Mat template, double threshold)
        {
            System.Drawing.Point? bestMatch = null;
            double bestScore = 0;
            double bestScale = 1.0;

            foreach (var scale in DefaultScales)
            {
                // テンプレートをスケーリング
                var newSize = new OpenCvSharp.Size(
                    (int)(template.Width * scale),
                    (int)(template.Height * scale)
                );

                // 画面サイズを超える場合はスキップ
                if (newSize.Width > screenshot.Width || newSize.Height > screenshot.Height)
                    continue;

                // 小さすぎる場合もスキップ
                if (newSize.Width < 10 || newSize.Height < 10)
                    continue;

                using var scaledTemplate = new Mat();
                Cv2.Resize(template, scaledTemplate, newSize, interpolation: InterpolationFlags.Cubic);

                // マッチング実行
                using var result = new Mat();
                Cv2.MatchTemplate(screenshot, scaledTemplate, result, TemplateMatchModes.CCoeffNormed);

                Cv2.MinMaxLoc(result, out _, out double maxVal, out _, out OpenCvSharp.Point maxLoc);

                if (maxVal > bestScore)
                {
                    bestScore = maxVal;
                    bestScale = scale;
                    // テンプレートの中心座標を計算
                    bestMatch = new System.Drawing.Point(
                        maxLoc.X + scaledTemplate.Width / 2,
                        maxLoc.Y + scaledTemplate.Height / 2
                    );
                }
            }

            if (bestScore >= threshold)
            {
                Logger.Instance.Info(
                    "ImageMatchingService",
                    $"画像マッチング成功: スコア={bestScore:F3}, スケール={bestScale:F2}x"
                );
                return bestMatch;
            }

            Logger.Instance.Warn(
                "ImageMatchingService",
                $"画像マッチング失敗: 最高スコア={bestScore:F3} (閾値={threshold:F3})"
            );
            return null;
        }

        /// <summary>
        /// 単一スケールマッチング
        /// </summary>
        private System.Drawing.Point? FindImageSingleScale(Mat screenshot, Mat template, double threshold)
        {
            using var result = new Mat();
            Cv2.MatchTemplate(screenshot, template, result, TemplateMatchModes.CCoeffNormed);

            Cv2.MinMaxLoc(result, out _, out double maxVal, out _, out OpenCvSharp.Point maxLoc);

            if (maxVal >= threshold)
            {
                Logger.Instance.Info("ImageMatchingService", $"画像マッチング成功: スコア={maxVal:F3}");

                return new System.Drawing.Point(
                    maxLoc.X + template.Width / 2,
                    maxLoc.Y + template.Height / 2
                );
            }

            Logger.Instance.Warn("ImageMatchingService", $"画像マッチング失敗: スコア={maxVal:F3} (閾値={threshold:F3})");
            return null;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }
    }

}

namespace RPACore
{
    /// <summary>
    /// NativeMethodsの拡張（スクリーンキャプチャ用）
    /// </summary>
    internal static partial class NativeMethods
    {
        [DllImport("user32.dll")]
        internal static extern int GetSystemMetrics(int nIndex);
    }
}
