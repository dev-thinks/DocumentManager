// ReSharper disable PossiblyMistakenUseOfParamsMethod
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Linq;
using DocumentManager.Core.MailMerge;
using HorizontalAnchorValues = DocumentFormat.OpenXml.Vml.Wordprocessing.HorizontalAnchorValues;
using Lock = DocumentFormat.OpenXml.Vml.Office.Lock;
using VerticalAnchorValues = DocumentFormat.OpenXml.Vml.Wordprocessing.VerticalAnchorValues;

namespace DocumentManager.Core.Converters.Handlers
{
    internal class DocxWatermark
    {
        private readonly ILogger _logger;
        private readonly WaterMarkOptions _options;
        private readonly MemoryStream _docxMs;

        private string WaterMarkTypeId => "#_x0000_t136";

        public DocxWatermark(string filePath, ILogger logger, WaterMarkOptions options)
        {
            _logger = logger;
            _options = options;
            _docxMs = Extensions.GetFileAsMemoryStream(filePath);

            if (options == null)
            {
                _options = new WaterMarkOptions();
            }
        }

        internal MemoryStream Do(string waterMarkImagePath = "")
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(_docxMs, true))
            {
                if (string.IsNullOrEmpty(waterMarkImagePath))
                {
                    _logger.LogTrace("Adding watermark text using: {WaterMarkImage}, {@Options}", waterMarkImagePath, _options);

                    AddWatermarkText(doc);
                }
                else
                {
                    _logger.LogTrace("Adding image watermark using: {WaterMarkImage}", waterMarkImagePath);
                    // TODO: Image watermark
                }
            }

            _docxMs.Position = 0;

            return _docxMs;
        }

        internal MemoryStream Remove()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(_docxMs, true))
            {
                RemoveWatermark(doc);
            }

            _docxMs.Position = 0;

            return _docxMs;
        }

        private void AddWatermarkText(WordprocessingDocument doc)
        {
            if (!doc.MainDocumentPart.HeaderParts.Any())
            {
                doc.MainDocumentPart.DeleteParts(doc.MainDocumentPart.HeaderParts);

                var sectionProps = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
                if (sectionProps == null)
                {
                    sectionProps = new SectionProperties();
                    doc.MainDocumentPart.Document.Body.Append(sectionProps);
                }

                HeaderPart headerPart1 = doc.MainDocumentPart.AddNewPart<HeaderPart>("rId7");
                OpenXmlDocxRef.GenerateHeaderPart1Content(headerPart1);
                var rId1 = doc.MainDocumentPart.GetIdOfPart(headerPart1);
                var headerRef1 = new HeaderReference { Id = rId1 };
                sectionProps.Append(headerRef1);

                HeaderPart headerPart2 = doc.MainDocumentPart.AddNewPart<HeaderPart>("rId6");
                OpenXmlDocxRef.GenerateHeaderPart2Content(headerPart2);
                var rId2 = doc.MainDocumentPart.GetIdOfPart(headerPart2);
                var headerRef2 = new HeaderReference { Id = rId2 };
                sectionProps.Append(headerRef2);

                HeaderPart headerPart3 = doc.MainDocumentPart.AddNewPart<HeaderPart>("rId10");
                OpenXmlDocxRef.GenerateHeaderPart3Content(headerPart3);
                var rId3 = doc.MainDocumentPart.GetIdOfPart(headerPart3);
                var headerRef3 = new HeaderReference { Id = rId3 };
                sectionProps.Append(headerRef3);
            }
        }

        private void RemoveWatermark(WordprocessingDocument doc)
        {
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                //Remove
                if (header.Header.Descendants<Paragraph>() != null)
                {
                    var isFound = false;
                    foreach (var para in header.Header.Descendants<Paragraph>())
                    {
                        foreach (Run r in para.Descendants<Run>())
                        {
                            isFound = FindAndRemoveWatermark(r);
                            if (isFound)
                                break;
                        }
                        if (isFound)
                            header.Header.Save(header);
                    }
                }
            }
        }

        private bool FindAndRemoveWatermark(Run runWatermark)
        {
            bool success = false;
            //DocumentFormat.OpenXml.Vml.TextPath
            //Check, if run contains watermark
            if (runWatermark.Descendants<Picture>() != null)
            {
                var listPic = runWatermark.Descendants<Picture>().ToList();

                for (int n = listPic.Count; n > 0; n--)
                {
                    if (listPic[n - 1].Descendants<Shape>() != null)
                    {
                        if (listPic[n - 1].Descendants<Shape>().Count(s => s.Type == WaterMarkTypeId) > 0)
                        {
                            //Found -> remove
                            listPic[n - 1].Remove();
                            success = true;
                            break;
                        }
                    }
                }
            }

            return success;
        }

        /*
        private void InsertCustomWatermark(WordprocessingDocument package, string p)
        {
            SetWaterMarkPicture(p);

            MainDocumentPart mainDocumentPart1 = package.MainDocumentPart;
            if (mainDocumentPart1 != null)
            {
                mainDocumentPart1.DeleteParts(mainDocumentPart1.HeaderParts);
                HeaderPart headPart1 = mainDocumentPart1.AddNewPart<HeaderPart>();
                GenerateHeaderPart1Content(headPart1);
                string rId = mainDocumentPart1.GetIdOfPart(headPart1);
                ImagePart image = headPart1.AddNewPart<ImagePart>("image/jpeg", "rId999");
                GenerateImagePart1Content(image);
                IEnumerable<SectionProperties> sectPrs = mainDocumentPart1.Document.Body.Elements<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    sectPr.RemoveAllChildren<HeaderReference>();
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() {Id = rId});
                }
            }
            else
            {

            }
        }

        private void SetWaterMarkPicture(string file)
        {
            FileStream inFile;
            byte[] byteArray;
            try
            {
                inFile = new FileStream(file, FileMode.Open, FileAccess.Read);
                byteArray = new byte[inFile.Length];
                long byteRead = inFile.Read(byteArray, 0, (int) inFile.Length);
                inFile.Close();
                _imagePart1Data = Convert.ToBase64String(byteArray, 0, byteArray.Length);
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }
        }

        private void GenerateHeaderPart1Content(HeaderPart headerPart1)
        {
            Header header1 = new Header();
            Paragraph paragraph2 = new Paragraph();
            Run run1 = new Run();
            Picture picture1 = new Picture();
            V.Shape shape1 = new V.Shape()
            {
                Id = "WordPictureWatermark75517470",
                Style =
                    "position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:415.2pt;height:456.15pt;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                OptionalString = "_x0000_s2051", AllowInCell = false, Type = "#_x0000_t75"
            };
            V.ImageData imageData1 = new V.ImageData()
                {Gain = "19661f", BlackLevel = "22938f", Title = "draft", RelationshipId = "rId999"};
            shape1.Append(imageData1);
            picture1.Append(shape1);
            run1.Append(picture1);
            paragraph2.Append(run1);
            header1.Append(paragraph2);
            headerPart1.Header = header1;
        }

        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(_imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        private static void DeleteCustomWatermark(WordprocessingDocument package, string watermarkId)
        {
            MainDocumentPart mainDocument = package.MainDocumentPart;

            var headers = mainDocument?.GetPartsOfType<HeaderPart>();
            if (headers != null)
            {
                foreach (var head in headers)
                {
                    var rt = mainDocument.GetIdOfPart(head);
                    if (string.Equals(rt, watermarkId, StringComparison.OrdinalIgnoreCase))
                    {
                        var watermark = head.GetPartById(watermarkId);
                        if (watermark != null)
                        {
                            head.DeletePart(watermark);
                        }
                    }
                }
            }
        }
        */
    }
}
