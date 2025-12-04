using System;
using System.IO;
using PdfSharpCore.Pdf;
using PdfSharpCore.Drawing;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    /// <summary>
    /// PDF生成器，每次截图时立即添加到PDF中
    /// </summary>
    public class PDFGenerator : IDisposable
    {
        private PdfDocument? _document;
        private readonly string _filePath;
        private int _pageWidth;
        private int _pageHeight;
        private bool _isInitialized = false;

        public PDFGenerator(string filePath, int width, int height)
        {
            _filePath = filePath;
            _pageWidth = width;
            _pageHeight = height;
        }

        /// <summary>
        /// 初始化PDF文件
        /// </summary>
        public void Initialize()
        {
            try
            {
                _document = new PdfDocument();
                _isInitialized = true;
                WriteLine($"PDF文件已初始化, 尺寸: {_pageWidth}x{_pageHeight}");
            }
            catch (Exception ex)
            {
                WriteError($"初始化PDF文件失败", ex);
                throw;
            }
        }

        /// <summary>
        /// 添加JPG图片到PDF
        /// </summary>
        public void AddImage(string imagePath)
        {
            try
            {
                if (!_isInitialized || _document == null)
                {
                    throw new InvalidOperationException("PDF未初始化");
                }

                // 创建新页面，尺寸与图片一致（转换为点，1英寸=72点，假设96 DPI）
                double widthPoints = _pageWidth * 72.0 / 96.0;
                double heightPoints = _pageHeight * 72.0 / 96.0;
                PdfPage page = _document.AddPage();
                page.Width = widthPoints;
                page.Height = heightPoints;

                // 添加图片到页面
                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                    XImage image = XImage.FromFile(imagePath);
                    gfx.DrawImage(image, 0, 0, widthPoints, heightPoints);
                    image.Dispose();
                }

                WriteLine($"已添加图片到PDF: {imagePath} (第 {_document.Pages.Count} 页)");
            }
            catch (Exception ex)
            {
                WriteError($"添加图片到PDF失败: {imagePath}", ex);
                throw;
            }
        }

        /// <summary>
        /// 完成PDF生成
        /// </summary>
        public void Finish()
        {
            try
            {
                if (_document != null)
                {
                    _document.Save(_filePath);
                    WriteLine($"PDF文件已生成完成, 共 {_document.Pages.Count} 页");
                }
            }
            catch (Exception ex)
            {
                WriteError($"完成PDF生成失败", ex);
                throw;
            }
        }

        public void Dispose()
        {
            try
            {
                _document?.Dispose();
            }
            catch (Exception ex)
            {
                WriteError($"释放PDF资源失败", ex);
            }
        }
    }
}

