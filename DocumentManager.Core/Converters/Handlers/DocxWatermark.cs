// ReSharper disable PossiblyMistakenUseOfParamsMethod
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Linq;
using HorizontalAnchorValues = DocumentFormat.OpenXml.Vml.Wordprocessing.HorizontalAnchorValues;
using Lock = DocumentFormat.OpenXml.Vml.Office.Lock;
using VerticalAnchorValues = DocumentFormat.OpenXml.Vml.Wordprocessing.VerticalAnchorValues;

namespace DocumentManager.Core.Converters.Handlers
{
    public class DocxWatermark
    {
        private readonly ILogger _logger;
        private readonly string _draftImagePath;
        private readonly MemoryStream _docxMs;

        public DocxWatermark(ILogger logger, string docXTemplateFilename, string draftImagePath = "")
        {
            _logger = logger;
            _draftImagePath = draftImagePath;
            _docxMs = StreamHandler.GetFileAsMemoryStream(docXTemplateFilename);
        }

        public MemoryStream Do()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(_docxMs, true))
            {
                AddWatermarkText(doc);
            }

            _docxMs.Position = 0;

            return _docxMs;
        }

        private static void AddWatermarkText(WordprocessingDocument doc)
        {
            if (!doc.MainDocumentPart.HeaderParts.Any())
            {
                doc.MainDocumentPart.DeleteParts(doc.MainDocumentPart.HeaderParts);
                var newHeaderPart = doc.MainDocumentPart.AddNewPart<HeaderPart>();
                var rId = doc.MainDocumentPart.GetIdOfPart(newHeaderPart);
                var headerRef = new HeaderReference {Id = rId};

                var sectionProps = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
                if (sectionProps == null)
                {
                    sectionProps = new SectionProperties();
                    doc.MainDocumentPart.Document.Body.Append(sectionProps);
                }

                sectionProps.RemoveAllChildren<HeaderReference>();
                sectionProps.Append(headerRef);

                newHeaderPart.Header = MakeHeader();
                newHeaderPart.Header.Save();
            }

            foreach (HeaderPart headerPart in doc.MainDocumentPart.HeaderParts)
            {
                var sdtBlock = new SdtBlock();
                var sdtProperties = new SdtProperties();
                var sdtId = new SdtId() {Val = 87908844};

                var sdtContentDocPartObject = new SdtContentDocPartObject();
                var docPartGallery = new DocPartGallery() {Val = "Watermarks"};
                var docPartUnique = new DocPartUnique();

                sdtContentDocPartObject.Append(docPartGallery);
                sdtContentDocPartObject.Append(docPartUnique);
                sdtProperties.Append(sdtId);
                sdtProperties.Append(sdtContentDocPartObject);

                var sdtContentBlock = new SdtContentBlock();
                var paragraph = new Paragraph()
                {
                    RsidParagraphAddition = "00656E18",
                    RsidRunAdditionDefault = "00656E18"
                };

                var paragraphProperties = new ParagraphProperties();
                var paragraphStyleId = new ParagraphStyleId() {Val = "Header"};
                paragraphProperties.Append(paragraphStyleId);

                var run1 = new Run();
                var runProperties = new RunProperties();
                var noProof = new NoProof();
                runProperties.Append(noProof);
                var picture = new Picture();
                var shapeType = new Shapetype
                {
                    Id = "_x0000_t136",
                    CoordinateSize = "21600,21600",
                    OptionalNumber = 136,
                    Adjustment = "10800",
                    EdgePath = "m@7,l@8,m@5,21600l@6,21600e"
                };

                var formulas = new Formulas();
                var formula1 = new Formula() {Equation = "sum #0 0 10800"};
                var formula2 = new Formula() {Equation = "prod #0 2 1"};
                var formula3 = new Formula() {Equation = "sum 21600 0 @1"};
                var formula4 = new Formula() {Equation = "sum 0 0 @2"};
                var formula5 = new Formula() {Equation = "sum 21600 0 @3"};
                var formula6 = new Formula() {Equation = "if @0 @3 0"};
                var formula7 = new Formula() {Equation = "if @0 21600 @1"};
                var formula8 = new Formula() {Equation = "if @0 0 @2"};
                var formula9 = new Formula() {Equation = "if @0 @4 21600"};
                var formula10 = new Formula() {Equation = "mid @5 @6"};
                var formula11 = new Formula() {Equation = "mid @8 @5"};
                var formula12 = new Formula() {Equation = "mid @7 @8"};
                var formula13 = new Formula() {Equation = "mid @6 @7"};
                var formula14 = new Formula() {Equation = "sum @6 0 @5"};

                formulas.Append(formula1);
                formulas.Append(formula2);
                formulas.Append(formula3);
                formulas.Append(formula4);
                formulas.Append(formula5);
                formulas.Append(formula6);
                formulas.Append(formula7);
                formulas.Append(formula8);
                formulas.Append(formula9);
                formulas.Append(formula10);
                formulas.Append(formula11);
                formulas.Append(formula12);
                formulas.Append(formula13);
                formulas.Append(formula14);

                var vmlPath = new DocumentFormat.OpenXml.Vml.Path()
                {
                    AllowTextPath = true,
                    ConnectionPointType = ConnectValues.Custom,
                    ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800",
                    ConnectAngles = "270,180,90,0"
                };
                var textPath = new TextPath()
                {
                    On = DocumentFormat.OpenXml.TrueFalseValue.FromBoolean(true),
                    FitShape = DocumentFormat.OpenXml.TrueFalseValue.FromBoolean(true)
                };

                var shapeHandles = new ShapeHandles();
                var shapeHandle = new ShapeHandle
                {
                    Position = "#0,bottomRight",
                    XRange = "6629,14971"
                };

                shapeHandles.Append(shapeHandle);

                var vmlLock = new Lock
                {
                    Extension = ExtensionHandlingBehaviorValues.Edit,
                    TextLock = DocumentFormat.OpenXml.TrueFalseValue.FromBoolean(true),
                    ShapeType = DocumentFormat.OpenXml.TrueFalseValue.FromBoolean(true)
                };

                shapeType.Append(formulas);
                shapeType.Append(vmlPath);
                shapeType.Append(textPath);
                shapeType.Append(shapeHandles);
                shapeType.Append(vmlLock);

                var vmlShape = new Shape()
                {
                    Id = "PowerPlusWaterMarkObject357476642",
                    Style =
                        "position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:315;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                    OptionalString = "_x0000_s2049",
                    AllowInCell = DocumentFormat.OpenXml.TrueFalseValue.FromBoolean(true),
                    FillColor = "silver",
                    Stroked = DocumentFormat.OpenXml.TrueFalseValue.FromBoolean(true),
                    Type = "#_x0000_t136"
                };

                var vmlFill = new Fill() {Opacity = ".5"};
                var fonTextPath = new TextPath()
                {
                    Style = "font-family:\"Calibri\";font-size:1pt",
                    String = "DRAFT"
                };

                var textWrap = new TextWrap()
                {
                    AnchorX = HorizontalAnchorValues.Margin,
                    AnchorY = VerticalAnchorValues.Margin
                };

                vmlShape.Append(vmlFill);
                vmlShape.Append(fonTextPath);
                vmlShape.Append(textWrap);
                
                picture.Append(shapeType);
                picture.Append(vmlShape);
                
                run1.Append(runProperties);
                run1.Append(picture);
                paragraph.Append(paragraphProperties);
                paragraph.Append(run1);
                sdtContentBlock.Append(paragraph);
                
                sdtBlock.Append(sdtProperties);
                sdtBlock.Append(sdtContentBlock);

                headerPart.Header.Append(sdtBlock);
                headerPart.Header.Save();
            }
        }

        private static Header MakeHeader()
        {
            var header = new Header();
            var paragraph = new Paragraph();
            var run = new Run();
            var text = new Text { Text = "" };

            run.Append(text);
            paragraph.Append(run);
            header.Append(paragraph);

            return header;
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
        */

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
