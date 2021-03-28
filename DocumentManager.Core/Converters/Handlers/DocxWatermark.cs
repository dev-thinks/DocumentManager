using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Logging;
using V = DocumentFormat.OpenXml.Vml;

namespace DocumentManager.Core.Converters.Handlers
{
    public class DocxWatermark
    {
        private readonly ILogger _logger;
        private readonly string _draftImagePath;
        private string _imagePart1Data = "";
        private readonly MemoryStream _docxMs;

        public DocxWatermark(ILogger logger, string docXTemplateFilename, string draftImagePath = "")
        {
            _logger = logger;
            _draftImagePath = draftImagePath;
            _docxMs = StreamHandler.GetFileAsMemoryStream(docXTemplateFilename);

            if (string.IsNullOrEmpty(draftImagePath))
            {
                _draftImagePath = "draft.jpg";
            }
        }

        public MemoryStream Do()
        {
            using (WordprocessingDocument package = WordprocessingDocument.Open(_docxMs, true))
            {
                InsertCustomWatermark(package, _draftImagePath);
            }

            _docxMs.Position = 0;

            return _docxMs;
        }

        Header MakeHeader()
        {
            var header = new Header();
            var paragraph = new Paragraph();
            var run = new Run();
            var text = new Text();
            text.Text = "";
            run.Append(text);
            paragraph.Append(run);
            header.Append(paragraph);
            return header;
        }

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
    }
}
