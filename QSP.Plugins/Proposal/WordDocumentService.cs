using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Xrm.Sdk;
using QSP.Plugins.Constants;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

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
            ProposalResult proposal)
        {
            if (templateBytes == null ||
                templateBytes.Length == 0)
            {
                throw new InvalidPluginExecutionException(
                    "Word template is empty.");
            }

            if (proposal == null)
            {
                throw new InvalidPluginExecutionException(
                    "Proposal data is empty.");
            }

            if (proposal.LogoBytes == null ||
                proposal.LogoBytes.Length == 0)
            {
                throw new InvalidPluginExecutionException(
                    "Business Unit logo is empty.");
            }

            _tracing.Trace(
                $"Generating Word document. Template size: {templateBytes.Length} bytes.");

            using (var stream =
                new MemoryStream())
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
                    var mainPart =
                        document.MainDocumentPart;

                    if (mainPart == null)
                    {
                        throw new InvalidPluginExecutionException(
                            "Word document does not contain a MainDocumentPart.");
                    }

                    ReplaceTextContentControl(
                        mainPart,
                        QspProposalConstants.ProposalDateTag,
                        proposal.ProposalDate);

                    ReplaceTextContentControl(
                        mainPart,
                        QspProposalConstants.QuoteNumberTag,
                        proposal.QuoteNumber);

                    ReplaceTextContentControl(
                        mainPart,
                        QspProposalConstants.ClientNameTag,
                        proposal.ClientName);

                    ReplaceTextContentControl(
                        mainPart,
                        QspProposalConstants.ContactNameTag,
                        proposal.ContactName);

                    ReplaceTextContentControl(
                        mainPart,
                        QspProposalConstants.ProjectNameTag,
                        proposal.ProjectName);

                    ReplaceTextContentControl(
                        mainPart,
                        QspProposalConstants.SiteAddressTag,
                        proposal.SiteAddress);

                    ReplaceTextContentControl(
                        mainPart,
                        QspProposalConstants.SellPriceTag,
                        proposal.SellPrice);

                    ReplaceBusinessUnitLogo(
                        mainPart,
                        proposal.LogoBytes);

                    mainPart.Document.Save();
                }

                var result =
                    stream.ToArray();

                _tracing.Trace(
                    $"Word document generated successfully. Size: {result.Length} bytes.");

                return result;
            }
        }

        private void ReplaceTextContentControl(
            MainDocumentPart mainPart,
            string tagName,
            string value)
        {
            value = value ?? string.Empty;

            _tracing.Trace(
                $"Looking for text control. Tag: {tagName}, Value: '{value}'");

            var allControls =
                mainPart.Document
                    .Descendants<SdtElement>()
                    .ToList();

            _tracing.Trace(
                $"Total content controls found: {allControls.Count}");

            foreach (var control in allControls)
            {
                var tag =
                    control.SdtProperties?
                        .GetFirstChild<Tag>()?
                        .Val?
                        .Value;

                if (!string.IsNullOrWhiteSpace(tag))
                {
                    _tracing.Trace(
                        $"Word content control found with tag: '{tag}'");
                }
            }

            var controls =
                allControls
                    .Where(control =>
                    {
                        var tag =
                            control.SdtProperties?
                                .GetFirstChild<Tag>()?
                                .Val?
                                .Value;

                        return string.Equals(
                            tag,
                            tagName,
                            StringComparison.Ordinal);
                    })
                    .ToList();

            if (controls.Count == 0)
            {
                throw new InvalidPluginExecutionException(
                    $"Word content control '{tagName}' was not found.");
            }

            _tracing.Trace(
                $"Found {controls.Count} control(s) for tag '{tagName}'. Value: '{value}'");

            foreach (var control in controls)
            {
                ReplaceTextContent(
                    control,
                    value);
            }

            _tracing.Trace(
                $"Finished processing content control '{tagName}'.");
        }

        private void ReplaceTextContent(
            SdtElement control,
            string value)
        {
            value = value ?? string.Empty;

            var textNodes =
                control
                    .Descendants<Text>()
                    .ToList();

            _tracing.Trace(
                $"Replacing content control text. Existing text nodes: {textNodes.Count}, New value: '{value}'");

            if (textNodes.Count == 0)
            {
                throw new InvalidPluginExecutionException(
                    "The Word content control does not contain a text element.");
            }

            textNodes[0].Text =
                value;

            textNodes[0].Space =
                SpaceProcessingModeValues.Preserve;

            for (int i = 1;
                i < textNodes.Count;
                i++)
            {
                textNodes[i].Text =
                    string.Empty;
            }

            _tracing.Trace(
                $"Content control text replaced successfully with '{value}'.");
        }

        private void ReplaceBusinessUnitLogo(
            MainDocumentPart mainPart,
            byte[] logoBytes)
        {
            _tracing.Trace(
                $"Searching for content control: {QspProposalConstants.LogoTag}");

            var logoControl =
                mainPart.Document
                    .Descendants<SdtElement>()
                    .FirstOrDefault(
                        control =>
                        {
                            var tag =
                                control.SdtProperties?
                                    .GetFirstChild<Tag>();

                            return tag != null &&
                                   string.Equals(
                                       tag.Val?.Value,
                                       QspProposalConstants.LogoTag,
                                       StringComparison.Ordinal);
                        });

            if (logoControl == null)
            {
                throw new InvalidPluginExecutionException(
                    $"Word content control '{QspProposalConstants.LogoTag}' was not found.");
            }

            _tracing.Trace(
                $"Content control '{QspProposalConstants.LogoTag}' found.");

            var imagePartType =
                GetImagePartType(
                    logoBytes);

            var imagePart =
                mainPart.AddImagePart(
                    imagePartType);

            using (var imageStream =
                new MemoryStream(
                    logoBytes))
            {
                imagePart.FeedData(
                    imageStream);
            }

            var relationshipId =
                mainPart.GetIdOfPart(
                    imagePart);

            var dimensions =
                CalculateLogoDimensions(
                    logoBytes);

            _tracing.Trace(
                $"Creating logo drawing. Width EMU: {dimensions.Item1}, Height EMU: {dimensions.Item2}");

            var drawing =
                CreateImageDrawing(
                    relationshipId,
                    dimensions.Item1,
                    dimensions.Item2);

            ReplaceContentControlWithImage(
                logoControl,
                drawing);

            _tracing.Trace(
                "Business Unit logo inserted successfully.");
        }

        private void ReplaceContentControlWithImage(
            SdtElement logoControl,
            Drawing drawing)
        {
            var paragraph =
                new Paragraph(
                    new ParagraphProperties(
                        new Justification
                        {
                            Val =
                                JustificationValues.Center
                        }),
                    new Run(
                        drawing));

            var parent =
                logoControl.Parent;

            if (parent == null)
            {
                throw new InvalidPluginExecutionException(
                    "Logo content control does not have a valid parent element.");
            }

            if (parent is Paragraph parentParagraph)
            {
                var grandParent =
                    parentParagraph.Parent;

                if (grandParent == null)
                {
                    throw new InvalidPluginExecutionException(
                        "Logo paragraph does not have a valid parent element.");
                }

                grandParent.InsertAfter(
                    paragraph,
                    parentParagraph);

                parentParagraph.Remove();

                _tracing.Trace(
                    "Logo Picture Content Control paragraph replaced.");
            }
            else
            {
                parent.InsertAfter(
                    paragraph,
                    logoControl);

                logoControl.Remove();

                _tracing.Trace(
                    "Logo Picture Content Control replaced.");
            }
        }

        private Drawing CreateImageDrawing(
            string relationshipId,
            long widthEmu,
            long heightEmu)
        {
            return new Drawing(
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
                            NoChangeAspect =
                                true
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
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip
                                    {
                                        Embed =
                                            relationshipId,
                                        CompressionState =
                                            A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(
                                        new A.FillRectangle())),
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
                                        }),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList())
                                    {
                                        Preset =
                                            A.ShapeTypeValues.Rectangle
                                    })))
                        {
                            Uri =
                                "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }))
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                });
        }

        private Tuple<long, long> CalculateLogoDimensions(
            byte[] logoBytes)
        {
            const long emuPerInch =
                914400L;

            const double maxWidthInches =
                6.5;

            const double maxHeightInches =
                1.5;

            using (var stream =
                new MemoryStream(
                    logoBytes))
            using (var image =
                Image.FromStream(
                    stream))
            {
                double pixelWidth =
                    image.Width;

                double pixelHeight =
                    image.Height;

                if (pixelWidth <= 0 ||
                    pixelHeight <= 0)
                {
                    throw new InvalidPluginExecutionException(
                        "Business Unit logo has invalid dimensions.");
                }

                double aspectRatio =
                    pixelWidth /
                    pixelHeight;

                double finalWidth =
                    maxWidthInches;

                double finalHeight =
                    finalWidth /
                    aspectRatio;

                if (finalHeight >
                    maxHeightInches)
                {
                    finalHeight =
                        maxHeightInches;

                    finalWidth =
                        finalHeight *
                        aspectRatio;
                }

                long widthEmu =
                    (long)(
                        finalWidth *
                        emuPerInch);

                long heightEmu =
                    (long)(
                        finalHeight *
                        emuPerInch);

                _tracing.Trace(
                    $"Original logo dimensions: {image.Width}x{image.Height}px");

                _tracing.Trace(
                    $"Logo aspect ratio: {aspectRatio:F2}");

                _tracing.Trace(
                    $"Final logo dimensions: {finalWidth:F2}x{finalHeight:F2} inches");

                return Tuple.Create(
                    widthEmu,
                    heightEmu);
            }
        }

        private PartTypeInfo GetImagePartType(
            byte[] imageBytes)
        {
            if (imageBytes == null ||
                imageBytes.Length < 4)
            {
                throw new InvalidPluginExecutionException(
                    "Business Unit logo is not a valid image.");
            }

            // PNG
            if (imageBytes.Length >= 8 &&
                imageBytes[0] == 0x89 &&
                imageBytes[1] == 0x50 &&
                imageBytes[2] == 0x4E &&
                imageBytes[3] == 0x47)
            {
                return ImagePartType.Png;
            }

            // JPEG
            if (imageBytes[0] == 0xFF &&
                imageBytes[1] == 0xD8)
            {
                return ImagePartType.Jpeg;
            }

            // GIF
            if (imageBytes[0] == 0x47 &&
                imageBytes[1] == 0x49 &&
                imageBytes[2] == 0x46)
            {
                return ImagePartType.Gif;
            }

            // BMP
            if (imageBytes[0] == 0x42 &&
                imageBytes[1] == 0x4D)
            {
                return ImagePartType.Bmp;
            }

            throw new InvalidPluginExecutionException(
                "Unsupported Business Unit logo image format. Use PNG, JPEG, GIF or BMP.");
        }
    }
}