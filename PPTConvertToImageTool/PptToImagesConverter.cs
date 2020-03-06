using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using PixelFormat = System.Drawing.Imaging.PixelFormat;
using Color = System.Drawing.Color;
using Size = System.Windows.Size;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using Microsoft.Office.Core;

namespace PPTConvertToImageTool
{
    /// <summary>
    /// Ppt转Enbx转换器。
    /// </summary>
    internal class PptToImagesConverter
    {
        /// <summary>
        /// 使用密码打开ppt（如果课件无密码则正常导入，密码错误则会抛密码错误异常，这里我们使用一个密码“PASSWORD”进行解密）；详见：https://stackoverflow.com/questions/17554892/unable-to-gracefully-abort-on-unknown-password-via-microsoft-office-interop-powe
        /// </summary>
        private const string PASSWORD_MARK = "::PASSWORD::";

        private const string ImageExtension = ".png";

        /// <summary>
        /// 允许导入的最大Ppt文件页数
        /// </summary>
        public static readonly int MaxPptSlidesCount = 99;

        /*
         * 默认截图的高宽 为 1280*720
         */
        private static readonly int DefaultWidth = 1280;

        private Size DefaultAspectRatio { get; }

        private double DefaultRatio { get; }

        private int DefaultHeight { get; }

        public PptToImagesConverter()
        {
            DefaultAspectRatio = new Size(16, 9);
            DefaultRatio = DefaultAspectRatio.Width / DefaultAspectRatio.Height;
            DefaultHeight = (int)(DefaultWidth / DefaultRatio);
        }

        /// <summary>
        /// 获取图片
        /// </summary>
        /// <param name="pptFile"></param>
        /// <param name="exportImagesFolder">导出图片目录</param>
        /// <returns></returns>
        public bool ConvertToImages(string pptFile, string exportImagesFolder)
        {
            try
            {
                var tempPpt = CopyTempPpt(pptFile);
                Application app = new Application();
                Presentation presentation = app.Presentations.Open(tempPpt + PASSWORD_MARK, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                if (presentation is null)
                {
                    Trace.WriteLine($"PPT文件打开失败,请检查PPT文件{pptFile}");
                    return false;
                }
                var images = new List<string>();
                var size = GetSlideSize(presentation, out bool isSizeChanged);
                var slides = GetPptSlide(presentation).Cast<Slide>();
                Parallel.ForEach(slides, slide =>
                {
                    string slidePath = Path.Combine(exportImagesFolder, "Slide-" + slide.SlideIndex + ImageExtension);
                    slide.Export(slidePath, ImageExtension, (int)size.Width, (int)size.Height);
                    if (File.Exists(slidePath))
                    {
                        lock (images)
                        {
                            images.Add(slidePath);
                        }
                    }
                });

                //调整图片
                AdjustImages(images, isSizeChanged);
                Dispose(app, presentation);
            }
            catch (Exception e)
            {
                Trace.WriteLine($"PPT导出失败{pptFile}，{e.Message}");
                return false;
            }
            return true;
        }

        private object CopyTempPpt(string file)
        {
            var tempPptPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + Path.GetExtension(file));
            File.Copy(file, tempPptPath);
            return tempPptPath;
        }

        /// <summary>
        /// 调整图片
        /// </summary>
        /// <param name="images"></param>
        /// <param name="isSizeChanged"></param>
        private void AdjustImages(List<string> images, bool isSizeChanged)
        {
            if (isSizeChanged)
            {
                ImageProcess(images, DefaultWidth, DefaultHeight);
            }
        }

        /// <summary>
        /// 获取Ppt页面
        /// </summary>
        /// <param name="presentationObject"></param>
        /// <returns></returns>
        private List<object> GetPptSlide(object presentationObject)
        {
            Presentation presentation = (Presentation)presentationObject;
            return presentation.Slides.Cast<object>().ToList();
        }

        /// <summary>
        /// 获取纵横比
        /// </summary>
        /// <param name="presentationObject"></param>
        /// <returns></returns>
        private float GetRatio(object presentationObject)
        {
            Presentation presentation = (Presentation)presentationObject;
            return presentation.PageSetup.SlideWidth / presentation.PageSetup.SlideHeight;
        }

        /// <summary>
        /// 获取页面尺寸
        /// </summary>
        /// <param name="presentationObject"></param>
        /// <param name="isSizeChanged"></param>
        /// <returns></returns>
        private Size GetSlideSize(object presentationObject, out bool isSizeChanged)
        {
            Presentation pp = (Presentation)presentationObject;
            var ratio = GetRatio(pp);
            var size = new Size(DefaultWidth, DefaultHeight);
            if (Math.Abs(DefaultRatio - ratio) < 0.001)
            {
                isSizeChanged = false;
            }
            else
            {
                if (DefaultRatio > ratio)
                {
                    //小于默认宽高比，则以高度为基准
                    size.Width = System.Convert.ToInt16(DefaultHeight * ratio);
                    size.Height = DefaultHeight;
                }
                else if (DefaultRatio < ratio)
                {
                    //小于默认宽高比，则以宽度为基准
                    size.Width = DefaultWidth;
                    size.Height = System.Convert.ToInt16(DefaultWidth / ratio);
                }
                isSizeChanged = true;
            }

            return size;
        }

        /// <summary>
        /// 清理资源
        /// </summary>
        /// <param name="applicationObject"></param>
        /// <param name="presentationObject"></param>
        private void Dispose(object applicationObject, object presentationObject)
        {
            Application app = (Application)applicationObject;
            Presentation pp = (Presentation)presentationObject;
            try
            {
                if (pp != null)
                {
                    pp.Close();
                    Marshal.ReleaseComObject(pp);
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
            }
            catch (Exception ex)
            {
                // 当 app or pp 带着异常进入时，这里可能再次抛出异常。
                // 如：上文中的 -2147467262 异常。
                Trace.WriteLine($"ppttoenbx:Error When Dispose. 异常信息可能是重复的-{ex.Message}");
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// 图片放大至标准比例并压缩处理
        /// </summary>
        /// <param name="files"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        private void ImageProcess(List<string> files, int width, int height)
        {
            if (files == null || !files.Any())
            {
                return;
            }
            string directoryName = new FileInfo(files.First()).DirectoryName;
            if (string.IsNullOrEmpty(directoryName))
            {
                return;
            }

            Parallel.ForEach(files, file => { AdjustImage(file, width, height); });
        }

        private void AdjustImage(string imageFilePath, int width, int height)
        {
            using (Bitmap newImage = new Bitmap(width, height, PixelFormat.Format24bppRgb))
            {
                using (Graphics g = Graphics.FromImage((System.Drawing.Image)newImage))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    using (Image image = Image.FromFile(imageFilePath))
                    {
                        int x = (width - image.Width) / 2;
                        int y = (height - image.Height) / 2;
                        g.DrawImage(image, x, y, image.Width, image.Height);
                    }
                }
                File.Delete(imageFilePath);
                newImage.Save(imageFilePath, ImageFormat.Png);
            }
        }
    }
}