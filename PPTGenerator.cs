using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    /// <summary>
    /// PPT 生成器：完全 C# 实现
    /// 1. 每次生成 PPT 时，在系统临时目录创建一个标准空白 PPT 模板
    /// 2. 用模板复制出目标 PPT，按截图尺寸设置页面，一张图一页
    /// 3. 生成完成后删除临时模板文件，用户只看到最终结果文件
    /// </summary>
    public class PPTGenerator : IDisposable
    {
        private PresentationDocument? _presentation;
        private PresentationPart? _presentationPart;
        private SlideMasterPart? _slideMasterPart;
        private SlideLayoutPart? _slideLayoutPart;

        private uint _slideId = 256;
        private uint _shapeId = 2;

        private readonly string _filePath;       // 最终输出的PPT路径
        private readonly int _slideWidth;        // 页面宽度（像素）
        private readonly int _slideHeight;       // 页面高度（像素）
        private readonly string _templatePath;   // 本次生成用的临时模板路径
        private bool _isInitialized = false;

        public PPTGenerator(string filePath, int width, int height)
        {
            _filePath = filePath;
            _slideWidth = width;
            _slideHeight = height;

            // 每次生成都在系统临时目录创建一个唯一模板文件
            _templatePath = Path.Combine(
                Path.GetTempPath(),
                $"ScreenshotPptTemplate_{Guid.NewGuid():N}.pptx");
        }

        /// <summary>
        /// 初始化 PPT：
        /// - 在临时目录创建模板
        /// - 复制模板为目标 PPT
        /// - 调整页面大小，清理模板内默认幻灯片
        /// </summary>
        public void Initialize()
        {
            try
            {
                if (_isInitialized) return;

                // 1. 如果目标文件已存在，删掉
                if (File.Exists(_filePath))
                {
                    File.Delete(_filePath);
                }

                // 2. 在临时目录创建一次性模板
                PptTemplateHelper.CreateTemplate(_templatePath);
                WriteLine($"已创建临时PPT模板: {_templatePath}");

                // 3. 从模板复制出本次要生成的 PPT 文件
                File.Copy(_templatePath, _filePath, overwrite: true);

                // 4. 打开复制出来的 PPT，做后续修改
                _presentation = PresentationDocument.Open(_filePath, true);
                _presentationPart = _presentation.PresentationPart
                    ?? throw new InvalidOperationException("模板缺少 PresentationPart");

                var presentation = _presentationPart.Presentation
                    ?? throw new InvalidOperationException("模板缺少 Presentation");

                // 5. 设置幻灯片尺寸：按截图像素换算成 inch → EMU
                double widthInches = Math.Max(_slideWidth / 96.0, 1.0);
                double heightInches = Math.Max(_slideHeight / 96.0, 1.0);

                presentation.SlideSize = new SlideSize
                {
                    Cx = (int)(widthInches * 914400),   // 1 inch = 914400 EMU
                    Cy = (int)(heightInches * 914400),
                    Type = SlideSizeValues.Custom
                };

                // 6. 拿到 Master / Layout，后面所有新建 Slide 都用这一套布局
                _slideMasterPart = _presentationPart.SlideMasterParts.FirstOrDefault()
                    ?? throw new InvalidOperationException("模板中没有 SlideMasterPart");

                _slideLayoutPart = _slideMasterPart.SlideLayoutParts.FirstOrDefault()
                    ?? throw new InvalidOperationException("模板中没有 SlideLayoutPart");

                // 7. 清空模板默认带的那一页幻灯片，只保留结构（Master/Layout/Theme）
                if (presentation.SlideIdList == null)
                {
                    presentation.SlideIdList = new SlideIdList();
                }
                else
                {
                    var slideIds = presentation.SlideIdList.Elements<SlideId>().ToList();
                    foreach (var slideId in slideIds)
                    {
                        string relId = slideId.RelationshipId;
                        var slidePart = (SlidePart)_presentationPart.GetPartById(relId);

                        presentation.SlideIdList.RemoveChild(slideId);
                        _presentationPart.DeletePart(slidePart);
                    }
                }

                // 从 256 开始给新 Slide 编号（与官方示例保持一致）
                _slideId = 256;
                _shapeId = 2;

                _isInitialized = true;
                WriteLine($"PPT文件已初始化: {_filePath}, 幻灯片页面尺寸(像素): {_slideWidth}x{_slideHeight}");
            }
            catch (Exception ex)
            {
                WriteError("初始化PPT文件失败", ex);
                // 初始化失败时也尝试删除临时模板
                TryDeleteTemplate();
                throw;
            }
        }

        /// <summary>
        /// 添加一张图片为一页幻灯片，按比例缩放并居中（参考 Python 版本逻辑）
        /// </summary>
        public void AddImage(string imagePath)
        {
            try
            {
                if (!_isInitialized || _presentationPart == null || _slideLayoutPart == null)
                    throw new InvalidOperationException("PPT未初始化");

                if (!File.Exists(imagePath))
                    throw new FileNotFoundException("图片文件不存在", imagePath);

                var presentation = _presentationPart.Presentation
                    ?? throw new InvalidOperationException("Presentation 不存在");

                // === 1. 新建 SlidePart，挂到当前 PPT 上 ===
                SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>();

                // Slide 本体：有一个 ShapeTree + ColorMapOverride（保证结构完整，参考官方示例）
                slidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = 1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new P.GroupShapeProperties(new D.TransformGroup())
                        )
                    ),
                    new ColorMapOverride(new D.MasterColorMapping())
                );

                // 关联到已有的 SlideLayout（所有页面使用同一个布局）
                slidePart.AddPart(_slideLayoutPart);

                // === 2. 添加图片部件 ===
                ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
                using (FileStream imageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    imagePart.FeedData(imageStream);
                }
                string imageRelId = slidePart.GetIdOfPart(imagePart);

                // === 3. 计算图片尺寸：直接铺满整个幻灯片 ===
                // 使用与PPT尺寸相同的计算方式，确保完全一致
                double widthInches = Math.Max(_slideWidth / 96.0, 1.0);
                double heightInches = Math.Max(_slideHeight / 96.0, 1.0);
                
                // 图片尺寸与PPT尺寸完全一致（使用相同的EMU计算方式）
                long picWidthEmu = (long)(widthInches * 914400);   // 1 inch = 914400 EMU
                long picHeightEmu = (long)(heightInches * 914400);
                long xEmu = 0;  // 从左上角开始
                long yEmu = 0;

                var picture = new P.Picture(
                    new P.NonVisualPictureProperties(
                        new P.NonVisualDrawingProperties
                        {
                            Id = _shapeId++,
                            Name = Path.GetFileName(imagePath)
                        },
                        new P.NonVisualPictureDrawingProperties(
                            new D.PictureLocks { NoChangeAspect = true }
                        ),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new P.BlipFill(
                        new D.Blip { Embed = imageRelId },
                        new D.Stretch(new D.FillRectangle())
                    ),
                    new P.ShapeProperties(
                        new D.Transform2D(
                            new D.Offset { X = xEmu, Y = yEmu },
                            new D.Extents { Cx = picWidthEmu, Cy = picHeightEmu }
                        ),
                        new D.PresetGeometry { Preset = D.ShapeTypeValues.Rectangle }
                    )
                );

                // 把图片塞进 Slide 的 ShapeTree 里
                slidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(picture);

                // === 4. 在 Presentation 的 SlideIdList 中登记这个 Slide ===
                if (presentation.SlideIdList == null)
                    presentation.SlideIdList = new SlideIdList();

                string relId = _presentationPart.GetIdOfPart(slidePart);

                presentation.SlideIdList.AppendChild(
                    new SlideId
                    {
                        Id = _slideId++,
                        RelationshipId = relId
                    }
                );

                slidePart.Slide.Save();

                WriteLine($"已添加图片到PPT: {imagePath} (当前总页数: {_slideId - 256})");
            }
            catch (Exception ex)
            {
                WriteError($"添加图片到PPT失败: {imagePath}", ex);
                throw;
            }
        }

        /// <summary>
        /// 完成 PPT 生成
        /// </summary>
        public void Finish()
        {
            try
            {
                if (!_isInitialized || _presentationPart == null || _presentation == null)
                    return;

                _presentationPart.Presentation.Save();
                _presentation.Dispose();
                _presentation = null;

                WriteLine($"PPT文件已生成完成: {_filePath}, 共 {_slideId - 256} 张幻灯片");
            }
            catch (Exception ex)
            {
                WriteError("完成PPT生成失败", ex);
                throw;
            }
            finally
            {
                _isInitialized = false;
                // 无论成功失败，都尝试删除模板
                TryDeleteTemplate();
            }
        }

        public void Dispose()
        {
            try
            {
                if (_presentation != null)
                {
                    _presentationPart?.Presentation?.Save();
                    _presentation.Dispose();
                    _presentation = null;
                }
            }
            catch (Exception ex)
            {
                WriteError("释放PPT资源失败", ex);
            }
            finally
            {
                // 再保险，Dispose 时也尝试删除模板
                TryDeleteTemplate();
            }
        }

        private void TryDeleteTemplate()
        {
            try
            {
                if (!string.IsNullOrEmpty(_templatePath) && File.Exists(_templatePath))
                {
                    File.Delete(_templatePath);
                    WriteLine($"已删除临时PPT模板: {_templatePath}");
                }
            }
            catch (Exception ex)
            {
                // 这里不要抛异常，只打印日志即可
                WriteError($"删除临时PPT模板失败: {_templatePath}", ex);
            }
        }
    }

    /// <summary>
    /// 模板生成帮助类：
    /// 每次被调用时，用 OpenXML 官方示例代码在指定路径创建一个标准空白 PPT。
    /// 这个模板是一次性的，用完就删。
    /// </summary>
    internal static class PptTemplateHelper
    {
        public static void CreateTemplate(string templatePath)
        {
            try
            {
                var dir = Path.GetDirectoryName(templatePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                if (File.Exists(templatePath))
                {
                    File.Delete(templatePath);
                }

                using (PresentationDocument presentationDoc =
                    PresentationDocument.Create(templatePath, PresentationDocumentType.Presentation))
                {
                    PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();

                    CreatePresentationParts(presentationPart);

                    presentationPart.Presentation.Save();
                }
            }
            catch (Exception ex)
            {
                WriteError("创建PPT模板失败", ex);
                throw;
            }
        }

        // 下面基本是 OpenXML 官方示例，用来创建一个结构完全规范的空白 PPT

        private static void CreatePresentationParts(PresentationPart presentationPart)
        {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(
                new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });

            SlideIdList slideIdList1 = new SlideIdList(
                new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });

            SlideSize slideSize1 = new SlideSize()
            {
                Cx = 9144000,
                Cy = 6858000,
                Type = SlideSizeValues.Screen4x3
            };

            NotesSize notesSize1 = new NotesSize()
            {
                Cx = 6858000,
                Cy = 9144000
            };

            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

            SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;

            slidePart1 = CreateSlidePart(presentationPart);
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
            themePart1 = CreateTheme(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");
        }

        private static SlidePart CreateSlidePart(PresentationPart presentationPart)
        {
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");

            slidePart1.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = 1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new P.GroupShapeProperties(new D.TransformGroup()),
                        new P.Shape(
                            new P.NonVisualShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = 2U, Name = "Title 1" },
                                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(new P.PlaceholderShape())
                            ),
                            new P.ShapeProperties(),
                            new P.TextBody(
                                new D.BodyProperties(),
                                new D.ListStyle(),
                                new D.Paragraph(new D.EndParagraphRunProperties() { Language = "en-US" })
                            )
                        )
                    )
                ),
                new ColorMapOverride(new D.MasterColorMapping())
            );

            return slidePart1;
        }

        private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
        {
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");

            SlideLayout slideLayout = new SlideLayout(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = 1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new P.GroupShapeProperties(new D.TransformGroup()),
                        new P.Shape(
                            new P.NonVisualShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = 2U, Name = "" },
                                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(new P.PlaceholderShape())
                            ),
                            new P.ShapeProperties(),
                            new P.TextBody(
                                new D.BodyProperties(),
                                new D.ListStyle(),
                                new D.Paragraph(new D.EndParagraphRunProperties())
                            )
                        )
                    )
                ),
                new ColorMapOverride(new D.MasterColorMapping())
            );

            slideLayoutPart1.SlideLayout = slideLayout;

            return slideLayoutPart1;
        }

        private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
        {
            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");

            SlideMaster slideMaster = new SlideMaster(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = 1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        ),
                        new P.GroupShapeProperties(new D.TransformGroup()),
                        new P.Shape(
                            new P.NonVisualShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = 2U, Name = "Title Placeholder 1" },
                                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(
                                    new P.PlaceholderShape() { Type = PlaceholderValues.Title })
                            ),
                            new P.ShapeProperties(),
                            new P.TextBody(
                                new D.BodyProperties(),
                                new D.ListStyle(),
                                new D.Paragraph()
                            )
                        )
                    )
                ),
                new P.ColorMap()
                {
                    Background1 = D.ColorSchemeIndexValues.Light1,
                    Text1 = D.ColorSchemeIndexValues.Dark1,
                    Background2 = D.ColorSchemeIndexValues.Light2,
                    Text2 = D.ColorSchemeIndexValues.Dark2,
                    Accent1 = D.ColorSchemeIndexValues.Accent1,
                    Accent2 = D.ColorSchemeIndexValues.Accent2,
                    Accent3 = D.ColorSchemeIndexValues.Accent3,
                    Accent4 = D.ColorSchemeIndexValues.Accent4,
                    Accent5 = D.ColorSchemeIndexValues.Accent5,
                    Accent6 = D.ColorSchemeIndexValues.Accent6,
                    Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                },
                new SlideLayoutIdList(
                    new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }
                ),
                new TextStyles(
                    new TitleStyle(),
                    new BodyStyle(),
                    new OtherStyle()
                )
            );

            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        private static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
        {
            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
            D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
                new D.ColorScheme(
                    new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
                    new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
                    new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
                    new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
                    new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
                    new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
                    new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
                    new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
                    new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
                    new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
                    new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" })
                )
                { Name = "Office" },
                new D.FontScheme(
                    new D.MajorFont(
                        new D.LatinFont() { Typeface = "Calibri" },
                        new D.EastAsianFont() { Typeface = "" },
                        new D.ComplexScriptFont() { Typeface = "" }
                    ),
                    new D.MinorFont(
                        new D.LatinFont() { Typeface = "Calibri" },
                        new D.EastAsianFont() { Typeface = "" },
                        new D.ComplexScriptFont() { Typeface = "" }
                    )
                )
                { Name = "Office" },
                new D.FormatScheme(
                    new D.FillStyleList(
                        new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                        new D.GradientFill(
                            new D.GradientStopList(
                                new D.GradientStop(
                                    new D.SchemeColor(
                                        new D.Tint() { Val = 50000 },
                                        new D.SaturationModulation() { Val = 300000 }
                                    )
                                    { Val = D.SchemeColorValues.PhColor }
                                ) { Position = 0 },
                                new D.GradientStop(
                                    new D.SchemeColor(
                                        new D.Tint() { Val = 37000 },
                                        new D.SaturationModulation() { Val = 300000 }
                                    )
                                    { Val = D.SchemeColorValues.PhColor }
                                ) { Position = 35000 },
                                new D.GradientStop(
                                    new D.SchemeColor(
                                        new D.Tint() { Val = 15000 },
                                        new D.SaturationModulation() { Val = 350000 }
                                    )
                                    { Val = D.SchemeColorValues.PhColor }
                                ) { Position = 100000 }
                            ),
                            new D.LinearGradientFill() { Angle = 16200000, Scaled = true }
                        ),
                        new D.NoFill(),
                        new D.PatternFill(),
                        new D.GroupFill()
                    ),
                    new D.LineStyleList(
                        new D.Outline(
                            new D.SolidFill(
                                new D.SchemeColor(
                                    new D.Shade() { Val = 95000 },
                                    new D.SaturationModulation() { Val = 105000 }
                                )
                                { Val = D.SchemeColorValues.PhColor }
                            ),
                            new D.PresetDash() { Val = D.PresetLineDashValues.Solid }
                        )
                        {
                            Width = 9525,
                            CapType = D.LineCapValues.Flat,
                            CompoundLineType = D.CompoundLineValues.Single,
                            Alignment = D.PenAlignmentValues.Center
                        },
                        new D.Outline(
                            new D.SolidFill(
                                new D.SchemeColor(
                                    new D.Shade() { Val = 95000 },
                                    new D.SaturationModulation() { Val = 105000 }
                                )
                                { Val = D.SchemeColorValues.PhColor }
                            ),
                            new D.PresetDash() { Val = D.PresetLineDashValues.Solid }
                        )
                        {
                            Width = 9525,
                            CapType = D.LineCapValues.Flat,
                            CompoundLineType = D.CompoundLineValues.Single,
                            Alignment = D.PenAlignmentValues.Center
                        },
                        new D.Outline(
                            new D.SolidFill(
                                new D.SchemeColor(
                                    new D.Shade() { Val = 95000 },
                                    new D.SaturationModulation() { Val = 105000 }
                                )
                                { Val = D.SchemeColorValues.PhColor }
                            ),
                            new D.PresetDash() { Val = D.PresetLineDashValues.Solid }
                        )
                        {
                            Width = 9525,
                            CapType = D.LineCapValues.Flat,
                            CompoundLineType = D.CompoundLineValues.Single,
                            Alignment = D.PenAlignmentValues.Center
                        }
                    ),
                    new D.EffectStyleList(
                        new D.EffectStyle(
                            new D.EffectList(
                                new D.OuterShadow(
                                    new D.RgbColorModelHex(
                                        new D.Alpha() { Val = 38000 }
                                    )
                                    { Val = "000000" }
                                )
                                {
                                    BlurRadius = 40000L,
                                    Distance = 20000L,
                                    Direction = 5400000,
                                    RotateWithShape = false
                                }
                            )
                        ),
                        new D.EffectStyle(
                            new D.EffectList(
                                new D.OuterShadow(
                                    new D.RgbColorModelHex(
                                        new D.Alpha() { Val = 38000 }
                                    )
                                    { Val = "000000" }
                                )
                                {
                                    BlurRadius = 40000L,
                                    Distance = 20000L,
                                    Direction = 5400000,
                                    RotateWithShape = false
                                }
                            )
                        ),
                        new D.EffectStyle(
                            new D.EffectList(
                                new D.OuterShadow(
                                    new D.RgbColorModelHex(
                                        new D.Alpha() { Val = 38000 }
                                    )
                                    { Val = "000000" }
                                )
                                {
                                    BlurRadius = 40000L,
                                    Distance = 20000L,
                                    Direction = 5400000,
                                    RotateWithShape = false
                                }
                            )
                        )
                    ),
                    new D.BackgroundFillStyleList(
                        new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                        new D.GradientFill(
                            new D.GradientStopList(
                                new D.GradientStop(
                                    new D.SchemeColor(
                                        new D.Tint() { Val = 50000 },
                                        new D.SaturationModulation() { Val = 300000 }
                                    )
                                    { Val = D.SchemeColorValues.PhColor }
                                ) { Position = 0 },
                                new D.GradientStop(
                                    new D.SchemeColor(
                                        new D.Tint() { Val = 50000 },
                                        new D.SaturationModulation() { Val = 300000 }
                                    )
                                    { Val = D.SchemeColorValues.PhColor }
                                ) { Position = 0 },
                                new D.GradientStop(
                                    new D.SchemeColor(
                                        new D.Tint() { Val = 50000 },
                                        new D.SaturationModulation() { Val = 300000 }
                                    )
                                    { Val = D.SchemeColorValues.PhColor }
                                ) { Position = 0 }
                            ),
                            new D.LinearGradientFill() { Angle = 16200000, Scaled = true }
                        ),
                        new D.GradientFill(
                            new D.GradientStopList(
                                new D.GradientStop(
                                    new D.SchemeColor(
                                        new D.Tint() { Val = 50000 },
                                        new D.SaturationModulation() { Val = 300000 }
                                    )
                                    { Val = D.SchemeColorValues.PhColor }
                                ) { Position = 0 },
                                new D.GradientStop(
                                    new D.SchemeColor(
                                        new D.Tint() { Val = 50000 },
                                        new D.SaturationModulation() { Val = 300000 }
                                    )
                                    { Val = D.SchemeColorValues.PhColor }
                                ) { Position = 0 }
                            ),
                            new D.LinearGradientFill() { Angle = 16200000, Scaled = true }
                        )
                    )
                )
                { Name = "Office" }
            );

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;

            return themePart1;
        }
    }
}
