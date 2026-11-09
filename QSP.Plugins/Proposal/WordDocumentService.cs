using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Microsoft.Xrm.Sdk;
using System;
using System.IO;
using System.Linq;
using System.Drawing;

namespace QSP.Plugins.Proposal
{
    internal class WordDocumentService
    {
        private readonly ITracingService _tracing;

        internal WordDocumentService(
            ITracingService tracing)
        {
            _tracing = tracing;
        }

        internal byte[] Generate(
            byte[] templateBytes,
            byte[] logoBytes)
        {
            if (templateBytes == null ||
                templateBytes.Length == 0)
            {
                throw new InvalidPluginExecutionException(
                    "Word template is empty.");
            }

            if (logoBytes == null ||
                logoBytes.Length == 0)
            {
                throw new InvalidPluginExecutionException(
                    "Business Unit logo is empty.");
            }

            using (var stream = new MemoryStream())
            {
                stream.Write(
                    templateBytes,
                    0,
                    templateBytes.Length);

                stream.Position = 0;

                using (var document =
                    WordprocessingDocument.Open(
                        stream,
                        true))
                {
                    _tracing.Trace(
                        "Word document opened.");

                    ReplaceLogo(
                        document,
                        logoBytes);

                    document.MainDocumentPart
                        .Document.Save();
                }

                _tracing.Trace(
                    "Logo replacement completed.");

                return stream.ToArray();
            }
        }

        private void ReplaceLogo(
            WordprocessingDocument document,
            byte[] logoBytes)
        {
            var mainPart =
                document.MainDocumentPart;

            var logoControl =
                mainPart.Document
                    .Descendants<SdtElement>()
                    .FirstOrDefault(x =>
                        x.SdtProperties?
                            .GetFirstChild<Tag>()?
                            .Val?.Value ==
                        Constants.QspProposalConstants.LogoTag);

            if (logoControl == null)
            {
                throw new InvalidPluginExecutionException(
                    "Word content control 'QSP_BU_LOGO' was not found.");
            }

            _tracing.Trace(
                "QSP_BU_LOGO content control found.");

            var imageType =
                GetImageType(logoBytes);

            var imagePart =
                mainPart.AddImagePart(imageType);

            using (var imageStream =
                new MemoryStream(logoBytes))
            {
                imagePart.FeedData(imageStream);
            }

            var relationshipId =
                mainPart.GetIdOfPart(imagePart);

            var drawing =
                CreateLogoDrawing(
                    relationshipId, logoBytes);

            ReplaceContentControlWithImage(
                logoControl,
                drawing);

            _tracing.Trace(
                "Business Unit logo inserted.");
        }

        private PartTypeInfo GetImageType(byte[] imageBytes)
        {
            if (imageBytes.Length >= 8 &&
                imageBytes[0] == 0x89 &&
                imageBytes[1] == 0x50 &&
                imageBytes[2] == 0x4E &&
                imageBytes[3] == 0x47)
            {
                return ImagePartType.Png;
            }

            if (imageBytes.Length >= 2 &&
                imageBytes[0] == 0xFF &&
                imageBytes[1] == 0xD8)
            {
                return ImagePartType.Jpeg;
            }

            throw new InvalidPluginExecutionException(
                "Business Unit logo must be PNG or JPEG.");
        }

        private Drawing CreateLogoDrawing(
            string relationshipId,
            byte[] logoBytes)
        {
            var dimensions =
                CalculateLogoDimensions(logoBytes);

            long widthEmu =
                dimensions.Item1;

            long heightEmu =
                dimensions.Item2;

            var drawing =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent
                        {
                            Cx = widthEmu,
                            Cy = heightEmu
                        },

                        new DW.EffectExtent
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },

                        new DW.DocProperties
                        {
                            Id = 1U,
                            Name = "QSP Business Unit Logo"
                        },

                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks
                            {
                                NoChangeAspect = true
                            }),

                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties
                                        {
                                            Id = 0U,
                                            Name = "QSP Business Unit Logo"
                                        },

                                        new PIC.NonVisualPictureDrawingProperties()
                                    ),

                                    new PIC.BlipFill(
                                        new A.Blip
                                        {
                                            Embed =
                                                relationshipId
                                        },

                                        new A.Stretch(
                                            new A.FillRectangle())
                                    ),

                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset
                                            {
                                                X = 0L,
                                                Y = 0L
                                            },

                                            new A.Extents
                                            {
                                                Cx = widthEmu,
                                                Cy = heightEmu
                                            }
                                        ),

                                        new A.PresetGeometry(
                                            new A.AdjustValueList())
                                        {
                                            Preset =
                                                A.ShapeTypeValues.Rectangle
                                        }
                                    )
                                )
                            )
                            {
                                Uri =
                                    "http://schemas.openxmlformats.org/drawingml/2006/picture"
                            }
                        )
                    )
                    {
                        DistanceFromTop = 0U,
                        DistanceFromBottom = 0U,
                        DistanceFromLeft = 0U,
                        DistanceFromRight = 0U
                    }
                );

            return drawing;
        }

        private Tuple<long, long> CalculateLogoDimensions(
    byte[] logoBytes)
        {
            const long emuPerInch = 914400L;

            const double maxWidthInches = 6.5;
            const double maxHeightInches = 1.5;

            using (var stream = new MemoryStream(logoBytes))
            using (var image = Image.FromStream(stream))
            {
                double pixelWidth = image.Width;
                double pixelHeight = image.Height;

                if (pixelWidth <= 0 || pixelHeight <= 0)
                {
                    throw new InvalidPluginExecutionException(
                        "Business Unit logo has invalid dimensions.");
                }

                double aspectRatio =
                    pixelWidth / pixelHeight;

                double finalWidth =
                    maxWidthInches;

                double finalHeight =
                    finalWidth / aspectRatio;

                if (finalHeight > maxHeightInches)
                {
                    finalHeight =
                        maxHeightInches;

                    finalWidth =
                        finalHeight * aspectRatio;
                }

                long widthEmu =
                    (long)(finalWidth * emuPerInch);

                long heightEmu =
                    (long)(finalHeight * emuPerInch);

                _tracing.Trace(
                    $"Original logo: {image.Width}x{image.Height}px");

                _tracing.Trace(
                    $"Logo aspect ratio: {aspectRatio:F2}");

                _tracing.Trace(
                    $"Final logo size: {finalWidth:F2}x{finalHeight:F2} inches");

                return Tuple.Create(
                    widthEmu,
                    heightEmu);
            }
        }

        private void ReplaceContentControlWithImage(
            SdtElement contentControl,
            Drawing drawing)
        {
            var content =
                contentControl.Descendants<SdtContentRun>()
                    .FirstOrDefault();

            if (content != null)
            {
                content.RemoveAllChildren();

                content.AppendChild(
                    new Run(drawing));

                return;
            }

            var blockContent =
                contentControl.Descendants<SdtContentBlock>()
                    .FirstOrDefault();

            if (blockContent != null)
            {
                blockContent.RemoveAllChildren();

                var paragraph =
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification
                            {
                                Val =
                                    JustificationValues.Center
                            }),
                        new Run(drawing));

                blockContent.AppendChild(paragraph);

                return;
            }

            throw new InvalidPluginExecutionException(
                "QSP_BU_LOGO has an unsupported Word content control structure.");
        }
    }
}