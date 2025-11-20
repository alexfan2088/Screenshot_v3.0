using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    /// <summary>
    /// PPT生成器，每次截图时立即添加到PPT中
    /// </summary>
    public class PPTGenerator : IDisposable
    {
        private PresentationDocument? _presentation;
        private PresentationPart? _presentationPart;
        private SlideMasterPart? _slideMasterPart;
        private SlideLayoutPart? _slideLayoutPart;
        private uint _slideId = 256;
        private uint _shapeId = 2;
        private readonly string _filePath;
        private int _slideWidth;
        private int _slideHeight;
        private bool _isInitialized = false;

        public PPTGenerator(string filePath, int width, int height)
        {
            _filePath = filePath;
            _slideWidth = width;
            _slideHeight = height;
        }

        /// <summary>
        /// 初始化PPT文件
        /// </summary>
        public void Initialize()
        {
            try
            {
                // 创建PPT文件
                _presentation = PresentationDocument.Create(_filePath, PresentationDocumentType.Presentation);
                _presentationPart = _presentation.AddPresentationPart();
                
                // 创建Presentation对象
                Presentation presentation = new Presentation();
                
                // 设置幻灯片尺寸（EMU单位：1英寸=914400 EMU）
                double widthInches = Math.Max(_slideWidth / 96.0, 1.0);
                double heightInches = Math.Max(_slideHeight / 96.0, 1.0);
                presentation.SlideSize = new SlideSize
                {
                    Cx = new Int32Value((int)(widthInches * 914400)),
                    Cy = new Int32Value((int)(heightInches * 914400)),
                    Type = SlideSizeValues.Custom
                };
                
                // 设置默认文本样式
                presentation.DefaultTextStyle = new DefaultTextStyle();
                
                // 创建幻灯片ID列表
                presentation.SlideIdList = new SlideIdList();
                
                // 创建幻灯片母版
                _slideMasterPart = _presentationPart.AddNewPart<SlideMasterPart>();
                SlideMaster slideMaster = new SlideMaster(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new GroupShapeProperties()
                        )
                    ),
                    new ColorMap(),
                    new TextStyles(
                        new TitleStyle(),
                        new BodyStyle(),
                        new OtherStyle()
                    )
                );
                _slideMasterPart.SlideMaster = slideMaster;
                
                // 创建幻灯片布局
                _slideLayoutPart = _slideMasterPart.AddNewPart<SlideLayoutPart>();
                SlideLayout slideLayout = new SlideLayout(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new GroupShapeProperties()
                        )
                    ),
                    new TextStyles(
                        new TitleStyle(),
                        new BodyStyle(),
                        new OtherStyle()
                    )
                );
                slideLayout.Type = SlideLayoutValues.Blank;
                _slideLayoutPart.SlideLayout = slideLayout;
                
                // 设置幻灯片母版关系
                SlideMasterIdList slideMasterIdList = new SlideMasterIdList();
                SlideMasterId slideMasterId = new SlideMasterId
                {
                    Id = 2147483648U,
                    RelationshipId = _presentationPart.GetIdOfPart(_slideMasterPart)
                };
                slideMasterIdList.Append(slideMasterId);
                presentation.SlideMasterIdList = slideMasterIdList;
                
                // 保存Presentation
                _presentationPart.Presentation = presentation;
                
                _isInitialized = true;
                WriteLine($"PPT文件已初始化: {_filePath}, 尺寸: {_slideWidth}x{_slideHeight}");
            }
            catch (Exception ex)
            {
                WriteError($"初始化PPT文件失败", ex);
                throw;
            }
        }

        /// <summary>
        /// 添加JPG图片到PPT
        /// </summary>
        public void AddImage(string imagePath)
        {
            try
            {
                if (!_isInitialized || _presentationPart == null || _slideLayoutPart == null)
                {
                    throw new InvalidOperationException("PPT未初始化");
                }

                // 创建新幻灯片
                SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>();
                
                // 设置幻灯片布局关系
                slidePart.AddPart(_slideLayoutPart);
                
                // 创建幻灯片（空白布局）
                Slide slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new GroupShapeProperties()
                        )
                    ),
                    new P.ColorMapOverride(
                        new A.MasterColorMapping()
                    )
                );
                slidePart.Slide = slide;
                
                // 添加图片部分
                ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
                using (FileStream imageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    imagePart.FeedData(imageStream);
                }

                // 获取图片关系ID
                string imageRelId = slidePart.GetIdOfPart(imagePart);
                
                // 计算图片尺寸（EMU单位：1像素 = 9525 EMU）
                long widthEmu = (long)(_slideWidth * 9525);
                long heightEmu = (long)(_slideHeight * 9525);

                // 创建图片形状
                var picture = new P.Picture(
                    new P.NonVisualPictureProperties(
                        new P.NonVisualDrawingProperties 
                        { 
                            Id = _shapeId++, 
                            Name = "图片" 
                        },
                        new P.NonVisualPictureDrawingProperties(
                            new A.PictureLocks { NoChangeAspect = true }
                        ),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new P.BlipFill(
                        new A.Blip 
                        { 
                            Embed = imageRelId,
                            CompressionState = A.BlipCompressionValues.Print
                        },
                        new A.Stretch(
                            new A.FillRectangle()
                        )
                    ),
                    new P.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = widthEmu, Cy = heightEmu }
                        ),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                    )
                );

                // 将图片添加到幻灯片
                if (slidePart.Slide?.CommonSlideData?.ShapeTree != null)
                {
                    slidePart.Slide.CommonSlideData.ShapeTree.AppendChild(picture);
                }

                // 添加幻灯片ID到列表
                string slideRelId = _presentationPart.GetIdOfPart(slidePart);
                if (_presentationPart.Presentation.SlideIdList != null)
                {
                    SlideId slideId = new SlideId
                    {
                        Id = _slideId++,
                        RelationshipId = slideRelId
                    };
                    _presentationPart.Presentation.SlideIdList.AppendChild(slideId);
                }

                WriteLine($"已添加图片到PPT: {imagePath} (幻灯片 {_slideId - 256})");
            }
            catch (Exception ex)
            {
                WriteError($"添加图片到PPT失败: {imagePath}", ex);
                throw;
            }
        }

        /// <summary>
        /// 完成PPT生成
        /// </summary>
        public void Finish()
        {
            try
            {
                if (_presentationPart != null && _presentationPart.Presentation != null)
                {
                    // 确保所有关系都已设置
                    if (_presentationPart.Presentation.SlideMasterIdList == null && _slideMasterPart != null)
                    {
                        SlideMasterIdList slideMasterIdList = new SlideMasterIdList();
                        SlideMasterId slideMasterId = new SlideMasterId
                        {
                            Id = 2147483648U,
                            RelationshipId = _presentationPart.GetIdOfPart(_slideMasterPart)
                        };
                        slideMasterIdList.Append(slideMasterId);
                        _presentationPart.Presentation.SlideMasterIdList = slideMasterIdList;
                    }
                }

                // 保存并关闭文档
                _presentationPart?.Presentation?.Save();
                _presentation?.Save();
                
                WriteLine($"PPT文件已生成完成: {_filePath}, 共 {_slideId - 256} 张幻灯片");
            }
            catch (Exception ex)
            {
                WriteError($"完成PPT生成失败", ex);
                throw;
            }
        }

        public void Dispose()
        {
            try
            {
                // 只释放资源，不调用Finish（Finish应该已经由外部调用）
                _presentation?.Dispose();
            }
            catch (Exception ex)
            {
                WriteError($"释放PPT资源失败", ex);
            }
        }
    }
}
