using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentManager.Core.Models;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace DocumentManager.Core.MailMerge
{
    public class OpenXmlDocxRef
    {
        // Generates content of headerPart1.
        public static void GenerateHeaderPart1Content(HeaderPart headerPart1, WaterMarkOptions options,
            string waterMarkTypeId)
        {
            Header header1 = new Header()
            {
                MCAttributes = new MarkupCompatibilityAttributes() {Ignorable = "w14 w15 w16se w16cid w16 w16cex wp14"}
            };
            header1.AddNamespaceDeclaration("wpc",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            header1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            header1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            header1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            header1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            header1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            header1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            header1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            header1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            header1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            header1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();
            SdtId sdtId1 = new SdtId() {Val = 1664587959};

            SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
            DocPartGallery docPartGallery1 = new DocPartGallery() {Val = "Watermarks"};
            DocPartUnique docPartUnique1 = new DocPartUnique();

            sdtContentDocPartObject1.Append(docPartGallery1);
            sdtContentDocPartObject1.Append(docPartUnique1);

            sdtProperties1.Append(sdtId1);
            sdtProperties1.Append(sdtContentDocPartObject1);

            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            Paragraph paragraph11 = new Paragraph()
            {
                RsidParagraphAddition = "00423092", RsidRunAdditionDefault = "00423092", ParagraphId = "789FD697",
                TextId = "2F71575D"
            };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() {Val = "Header"};

            paragraphProperties2.Append(paragraphStyleId2);

            Run run15 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            Picture picture1 = new Picture() {AnchorId = "699E31DA"};

            V.Shapetype shapetype1 = new V.Shapetype()
            {
                Id = waterMarkTypeId, CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800",
                EdgePath = "m@7,l@8,m@5,21600l@6,21600e"
            };

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula() {Equation = "sum #0 0 10800"};
            V.Formula formula2 = new V.Formula() {Equation = "prod #0 2 1"};
            V.Formula formula3 = new V.Formula() {Equation = "sum 21600 0 @1"};
            V.Formula formula4 = new V.Formula() {Equation = "sum 0 0 @2"};
            V.Formula formula5 = new V.Formula() {Equation = "sum 21600 0 @3"};
            V.Formula formula6 = new V.Formula() {Equation = "if @0 @3 0"};
            V.Formula formula7 = new V.Formula() {Equation = "if @0 21600 @1"};
            V.Formula formula8 = new V.Formula() {Equation = "if @0 0 @2"};
            V.Formula formula9 = new V.Formula() {Equation = "if @0 @4 21600"};
            V.Formula formula10 = new V.Formula() {Equation = "mid @5 @6"};
            V.Formula formula11 = new V.Formula() {Equation = "mid @8 @5"};
            V.Formula formula12 = new V.Formula() {Equation = "mid @7 @8"};
            V.Formula formula13 = new V.Formula() {Equation = "mid @6 @7"};
            V.Formula formula14 = new V.Formula() {Equation = "sum @6 0 @5"};

            formulas1.Append(formula1);
            formulas1.Append(formula2);
            formulas1.Append(formula3);
            formulas1.Append(formula4);
            formulas1.Append(formula5);
            formulas1.Append(formula6);
            formulas1.Append(formula7);
            formulas1.Append(formula8);
            formulas1.Append(formula9);
            formulas1.Append(formula10);
            formulas1.Append(formula11);
            formulas1.Append(formula12);
            formulas1.Append(formula13);
            formulas1.Append(formula14);
            V.Path path1 = new V.Path()
            {
                AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom,
                ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0"
            };
            V.TextPath textPath1 = new V.TextPath() {On = true, FitShape = true};

            V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
            V.ShapeHandle shapeHandle1 = new V.ShapeHandle()
                {Position = $"#0,{options.Position}", XRange = "6629,14971"};

            shapeHandles1.Append(shapeHandle1);
            Ovml.Lock lock1 = new Ovml.Lock()
                {Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true};

            shapetype1.Append(formulas1);
            shapetype1.Append(path1);
            shapetype1.Append(textPath1);
            shapetype1.Append(shapeHandles1);
            shapetype1.Append(lock1);

            V.Shape shape1 = new V.Shape()
            {
                Id = "PowerPlusWaterMarkObject357476642",
                Style = options.ElementStyle, OptionalString = "_x0000_s2049", AllowInCell = false,
                FillColor = options.ElementColor, Stroked = false, Type = $"#{waterMarkTypeId}"
            };
            V.Fill fill1 = new V.Fill() {Opacity = ".5"};
            V.TextPath textPath2 = new V.TextPath() {Style = options.ElementFontFamily, String = options.Text};
            Wvml.TextWrap textWrap1 = new Wvml.TextWrap()
                {AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin};

            shape1.Append(fill1);
            shape1.Append(textPath2);
            shape1.Append(textWrap1);

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            run15.Append(runProperties1);
            run15.Append(picture1);

            paragraph11.Append(paragraphProperties2);
            paragraph11.Append(run15);

            sdtContentBlock1.Append(paragraph11);

            sdtBlock1.Append(sdtProperties1);
            sdtBlock1.Append(sdtContentBlock1);

            header1.Append(sdtBlock1);

            headerPart1.Header = header1;
        }

        // Generates content of headerPart2.
        public static void GenerateHeaderPart2Content(HeaderPart headerPart2)
        {
            Header header2 = new Header()
            {
                MCAttributes = new MarkupCompatibilityAttributes() {Ignorable = "w14 w15 w16se w16cid w16 w16cex wp14"}
            };
            header2.AddNamespaceDeclaration("wpc",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header2.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header2.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header2.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            header2.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            header2.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            header2.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            header2.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            header2.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            header2.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            header2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header2.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            header2.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            header2.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header2.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header2.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header2.AddNamespaceDeclaration("wp14",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header2.AddNamespaceDeclaration("wp",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header2.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header2.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header2.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header2.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            header2.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            header2.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            header2.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header2.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header2.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header2.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header2.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph12 = new Paragraph()
            {
                RsidParagraphAddition = "00423092", RsidRunAdditionDefault = "00423092", ParagraphId = "5E3BA7DF",
                TextId = "77777777"
            };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() {Val = "Header"};

            paragraphProperties3.Append(paragraphStyleId3);

            paragraph12.Append(paragraphProperties3);

            header2.Append(paragraph12);

            headerPart2.Header = header2;
        }

        // Generates content of headerPart3.
        public static void GenerateHeaderPart3Content(HeaderPart headerPart3)
        {
            Header header3 = new Header()
            {
                MCAttributes = new MarkupCompatibilityAttributes() {Ignorable = "w14 w15 w16se w16cid w16 w16cex wp14"}
            };
            header3.AddNamespaceDeclaration("wpc",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header3.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header3.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header3.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            header3.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            header3.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            header3.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            header3.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            header3.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            header3.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            header3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header3.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            header3.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            header3.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header3.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header3.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header3.AddNamespaceDeclaration("wp14",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header3.AddNamespaceDeclaration("wp",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header3.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header3.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header3.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header3.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            header3.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            header3.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            header3.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header3.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header3.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header3.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header3.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph16 = new Paragraph()
            {
                RsidParagraphAddition = "00423092", RsidRunAdditionDefault = "00423092", ParagraphId = "21D8D7A9",
                TextId = "77777777"
            };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() {Val = "Header"};

            paragraphProperties7.Append(paragraphStyleId5);

            paragraph16.Append(paragraphProperties7);

            header3.Append(paragraph16);

            headerPart3.Header = header3;
        }

        public void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "KayKay";
            document.PackageProperties.Title = "";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.Revision = "3";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2021-03-29T18:46:00Z",
                System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2021-03-29T18:47:00Z",
                System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "KayKay";
        }

        public static void GenerateHeaderPartContent(HeaderPart headerPart1, WaterMarkOptions options,
            string waterMarkTypeId)
        {
            Header header1 = new Header() {MCAttributes = new MarkupCompatibilityAttributes() {Ignorable = "w14 wp14"}};
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("wp",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("wp14",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            Paragraph paragraph2 = new Paragraph();

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() {Val = "Header"};
            SuppressLineNumbers suppressLineNumbers5 = new SuppressLineNumbers();
            BiDi biDi3 = new BiDi() {Val = false};
            Justification justification2 = new Justification() {Val = JustificationValues.Left};
            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();

            paragraphProperties2.Append(paragraphStyleId2);
            paragraphProperties2.Append(suppressLineNumbers5);
            paragraphProperties2.Append(biDi3);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run();
            RunProperties runProperties2 = new RunProperties();

            Picture picture1 = new Picture();

            V.Shapetype shapetype1 = new V.Shapetype()
            {
                Id = waterMarkTypeId, CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800",
                EdgePath = "m@9,l@10,em@11,21600l@12,21600e"
            };
            V.Stroke stroke1 = new V.Stroke() {JoinStyle = V.StrokeJoinStyleValues.Miter};

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula() {Equation = "val #0"};
            V.Formula formula2 = new V.Formula() {Equation = "sum @0 0 10800"};
            V.Formula formula3 = new V.Formula() {Equation = "sum @0 0 0"};
            V.Formula formula4 = new V.Formula() {Equation = "sum width 0 @0"};
            V.Formula formula5 = new V.Formula() {Equation = "prod @2 2 1"};
            V.Formula formula6 = new V.Formula() {Equation = "prod @3 2 1"};
            V.Formula formula7 = new V.Formula() {Equation = "if @1 @5 @4"};
            V.Formula formula8 = new V.Formula() {Equation = "sum 0 @6 0"};
            V.Formula formula9 = new V.Formula() {Equation = "sum width 0 @6"};
            V.Formula formula10 = new V.Formula() {Equation = "if @1 0 @8"};
            V.Formula formula11 = new V.Formula() {Equation = "if @1 @7 width"};
            V.Formula formula12 = new V.Formula() {Equation = "if @1 @8 0"};
            V.Formula formula13 = new V.Formula() {Equation = "if @1 width @7"};

            formulas1.Append(formula1);
            formulas1.Append(formula2);
            formulas1.Append(formula3);
            formulas1.Append(formula4);
            formulas1.Append(formula5);
            formulas1.Append(formula6);
            formulas1.Append(formula7);
            formulas1.Append(formula8);
            formulas1.Append(formula9);
            formulas1.Append(formula10);
            formulas1.Append(formula11);
            formulas1.Append(formula12);
            formulas1.Append(formula13);

            V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
            V.ShapeHandle shapeHandle1 = new V.ShapeHandle() {Position = "@0,21600"};

            shapeHandles1.Append(shapeHandle1);

            shapetype1.Append(stroke1);
            shapetype1.Append(formulas1);
            shapetype1.Append(shapeHandles1);

            V.Shape shape1 = new V.Shape()
            {
                Id = "PowerPlusWaterMarkObject",
                Style =
                    "position:absolute;margin-left:0.05pt;margin-top:256.9pt;width:498.5pt;height:150.4pt;mso-wrap-style:none;v-text-anchor:middle;rotation:315;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                OptionalString = "shape_0", FillColor = options.ElementColor, Stroked = false, Type = waterMarkTypeId
            };
            V.Path path1 = new V.Path() {AllowTextPath = true};
            V.TextPath textPath1 = new V.TextPath()
            {
                Style = options.ElementFontFamily, On = true, FitShape = true, Trim = true,
                String = options.Text
            };
            V.Fill fill1 = new V.Fill()
                {Type = V.FillTypeValues.Solid, Opacity = "0.5", Color2 = "#3f3f3f", DetectMouseClick = true};
            V.Stroke stroke2 = new V.Stroke()
                {Color = "#3465a4", JoinStyle = V.StrokeJoinStyleValues.Round, EndCap = V.StrokeEndCapValues.Flat};
            Wvml.TextWrap textWrap1 = new Wvml.TextWrap() {Type = Wvml.WrapValues.None};

            shape1.Append(path1);
            shape1.Append(textPath1);
            shape1.Append(fill1);
            shape1.Append(stroke2);
            shape1.Append(textWrap1);

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            run2.Append(runProperties2);
            run2.Append(picture1);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            header1.Append(paragraph2);

            headerPart1.Header = header1;
        }
    }
}
