using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocumentManager.Core.Converters.Handlers
{
    public class ImageHandler
    {
        private readonly ILogger _logger;

        public ImageHandler(ILogger logger)
        {
            _logger = logger;
        }

        public void AppendImageToElement(KeyValuePair<string, ImageElement> placeholder, OpenXmlElement element,
            WordprocessingDocument doc, int imageCounter)
        {
            string imageExtension = placeholder.Value.MemStream.GetImageType();

            MainDocumentPart mainPart = doc.MainDocumentPart;

            var imageUri = new Uri($"/word/media/{placeholder.Key}{imageCounter}.{imageExtension}", UriKind.Relative);

            // Create "image" part in /word/media
            // Change content type for other image types.
            PackagePart packageImagePart = doc.Package.CreatePart(imageUri, "Image/" + imageExtension);

            // Feed data.
            placeholder.Value.MemStream.Position = 0;
            byte[] imageBytes = placeholder.Value.MemStream.ToArray();
            packageImagePart.GetStream().Write(imageBytes, 0, imageBytes.Length);

            PackagePart documentPackagePart =
                mainPart.OpenXmlPackage.Package.GetPart(new Uri("/word/document.xml", UriKind.Relative));

            // URI to the image is relative to relationship document.
            PackageRelationship imageRelationshipPart = documentPackagePart.CreateRelationship(
                new Uri("media/" + placeholder.Key + imageCounter + "." + imageExtension, UriKind.Relative),
                TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

            var imgTmp = placeholder.Value.MemStream.GetImage();

            var drawing = GetImageElement(imageRelationshipPart.Id, placeholder.Key, "picture", imgTmp.Width,
                imgTmp.Height, placeholder.Value.Dpi, imageCounter);
            element.AppendChild(drawing);
        }

        private Drawing GetImageElement(string imagePartId, string fileName, string pictureName,
            double width, double height, double ppi, int imageCounter)
        {
            double englishMetricUnitsPerInch = 914400;
            double pixelsPerInch = ppi;

            //calculate size in emu
            double emuWidth = width * englishMetricUnitsPerInch / pixelsPerInch;
            double emuHeight = height * englishMetricUnitsPerInch / pixelsPerInch;

            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent {Cx = (Int64Value) emuWidth, Cy = (Int64Value) emuHeight},
                    new DW.EffectExtent {LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L},
                    new DW.DocProperties {Id = (UInt32Value) 1U, Name = pictureName + imageCounter},
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks {NoChangeAspect = true}),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties {Id = (UInt32Value) 0U, Name = fileName},
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension {Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"}))
                                    {
                                        Embed = imagePartId,
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset {X = 0L, Y = 0L},
                                        new A.Extents {Cx = (Int64Value) emuWidth, Cy = (Int64Value) emuHeight}),
                                    new A.PresetGeometry(
                                            new A.AdjustValueList())
                                        {Preset = A.ShapeTypeValues.Rectangle})))
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }))
                {
                    DistanceFromTop = (UInt32Value) 0U,
                    DistanceFromBottom = (UInt32Value) 0U,
                    DistanceFromLeft = (UInt32Value) 0U,
                    DistanceFromRight = (UInt32Value) 0U,
                    EditId = "50D07946"
                });

            return element;
        }
    }
}
