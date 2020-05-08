using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Kybs0.Net.Utils;

namespace PPTConvertToImageTool
{
    public class PptToImageConverterByAspose
    {
        private const string ImageExtension = ".png";
        private const string SlideString = "Slide-";
        //默认截图的高宽 为 1280*720
        private static readonly int DefaultWidth = 1280;

        private Size DefaultAspectRatio { get; }

        private int DefaultHeight { get; }

        public PptToImageConverterByAspose()
        {
            DefaultAspectRatio = new Size(16, 9);
            var defaultRatio = DefaultAspectRatio.Width / Convert.ToDouble(DefaultAspectRatio.Height);
            DefaultHeight = (int)(DefaultWidth / defaultRatio);
        }

        public bool ConvertToImages(string pptFile, string exportImagesFolder)
        {
            try
            {
                if (!Directory.Exists(exportImagesFolder))
                {
                    Directory.CreateDirectory(exportImagesFolder);
                }
                var images = new List<string>();
                using (Presentation pres = new Presentation(pptFile))
                {
                    var slideSize = GetSlideSize(pres);
                    float scaleX = (float)((1.0 / pres.SlideSize.Size.Width) * slideSize.Width);
                    float scaleY = (float)((1.0 / pres.SlideSize.Size.Height) * slideSize.Height);
                    foreach (ISlide sld in pres.Slides)
                    {
                        Bitmap bmp = sld.GetThumbnail(scaleX, scaleY);
                        string slidePath = Path.Combine(exportImagesFolder, $"{SlideString}{sld.SlideNumber}{ImageExtension}");
                        if (File.Exists(slidePath))
                        {
                            File.Delete(slidePath);
                        }
                        bmp.Save(slidePath, ImageFormat.Png);
                        images.Add(slidePath);
                    }
                }
                //调整图片
                AdjustImages(images);

            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 调整图片
        /// </summary>
        /// <param name="images"></param>
        private void AdjustImages(List<string> images)
        {
            if (images == null || !images.Any())
            {
                return;
            }
            string directoryName = new FileInfo(images.First()).DirectoryName;
            if (string.IsNullOrEmpty(directoryName))
            {
                return;
            }
            Parallel.ForEach(images, file => { ImageSizeAdjustHelper.AdjustImageByMaxSize(file, DefaultWidth, DefaultHeight); });
        }

        /// <summary>
        /// 获取页面尺寸
        /// </summary>
        /// <param name="presentationObject"></param>
        /// <returns></returns>
        private System.Windows.Size GetSlideSize(object presentationObject)
        {
            Presentation presentation = (Presentation)presentationObject;
            var ratio = presentation.SlideSize.Size.Width / presentation.SlideSize.Size.Height;
            var size = new System.Windows.Size(DefaultWidth, DefaultHeight);
            var defaultRatio = DefaultAspectRatio.Width / Convert.ToDouble(DefaultAspectRatio.Height);
            if (Math.Abs(defaultRatio - ratio) < 0.001)
            {
            }
            else
            {
                if (defaultRatio > ratio)
                {
                    //小于默认宽高比，则以高度为基准
                    size.Width = System.Convert.ToInt16(DefaultHeight * ratio);
                    size.Height = DefaultHeight;
                }
                else if (defaultRatio < ratio)
                {
                    //小于默认宽高比，则以宽度为基准
                    size.Width = DefaultWidth;
                    size.Height = System.Convert.ToInt16(DefaultWidth / ratio);
                }
            }

            return size;
        }
    }
}
