// ReSharper disable PossiblyMistakenUseOfParamsMethod
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using A = DocumentFormat.OpenXml.Drawing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace DocumentManager.Core.Converters.Handlers
{
    internal class DocxStamp
    {
        private readonly ILogger _logger;
        private readonly StampMarkOptions _options;
        private readonly MemoryStream _docxMs;

        public DocxStamp(string filePath, ILogger logger, StampMarkOptions options)
        {
            _logger = logger;
            _options = options;
            _docxMs = Extensions.GetFileAsMemoryStream(filePath);

            if (options == null)
            {
                _options = new StampMarkOptions();
            }
        }

        internal MemoryStream Do()
        {
            // REMARKS!! - any id reference needs to be start from 200. hopefully it will not coincide with other elements in the document
            using (WordprocessingDocument document = WordprocessingDocument.Open(_docxMs, true))
            {
                ImagePart imagePart1 = document.MainDocumentPart.AddNewPart<ImagePart>("image/png", "rId200");
                GenerateImagePart1Content(imagePart1);

                var p = GenerateParagraph();
                document.MainDocumentPart.Document.Body.Append(p);

                var r = document.MainDocumentPart.Document.Body.OuterXml;
            }

            _docxMs.Position = 0;

            return _docxMs;
        }

        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        private string imagePart1Data =
            "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAIAQMAAAD+wSzIAAAABlBMVEX///8AAABVwtN+AAAAEklEQVQI12NgYHBhYGAQZIDSAAWCAKvnQf5VAAAAAElFTkSuQmCC";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        private Paragraph GenerateParagraph()
        {
            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() {Val = "Normal"};
            BiDi biDi1 = new BiDi() {Val = false};
            Justification justification1 = new Justification() {Val = JustificationValues.Left};

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() {Ascii = "Calibri", HighAnsi = "Calibri"};

            paragraphMarkRunProperties1.Append(runFonts1);

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(biDi1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();
            RunProperties runProperties1 = new RunProperties();

            AlternateContent alternateContent1 = new AlternateContent();

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() {Requires = "wps"};

            Drawing drawing1 = new Drawing();

            Wp.Anchor anchor1 = new Wp.Anchor()
            {
                DistanceFromTop = (UInt32Value) 0U, DistanceFromBottom = (UInt32Value) 0U,
                DistanceFromLeft = (UInt32Value) 0U, DistanceFromRight = (UInt32Value) 0U, SimplePos = false,
                RelativeHeight = (UInt32Value) 2U, BehindDoc = false, Locked = false, LayoutInCell = false,
                AllowOverlap = true
            };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() {X = 0L, Y = 0L};

            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition()
                {RelativeFrom = Wp.HorizontalRelativePositionValues.Column};
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = "4406900";

            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition()
                {RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph};
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "5220970";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent1 = new Wp.Extent() {Cx = 1926590L, Cy = 483235L};
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent()
                {LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L};
            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties1 = new Wp.DocProperties() {Id = (UInt32Value) 1U, Name = "Shape1"};

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData()
                {Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"};

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 =
                new Wps.NonVisualDrawingShapeProperties();

            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() {X = 0L, Y = 0L};
            A.Extents extents1 = new A.Extents() {Cx = 1926000L, Cy = 482760L};

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() {Preset = A.ShapeTypeValues.Rectangle};
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            A.BlipFill blipFill1 = new A.BlipFill() {RotateWithShape = false};
            A.Blip blip1 = new A.Blip() {Embed = "rId200"};
            A.Tile tile1 = new A.Tile();

            blipFill1.Append(blip1);
            blipFill1.Append(tile1);

            A.Outline outline1 = new A.Outline() {Width = 38160};

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() {Val = "000000"};

            solidFill1.Append(rgbColorModelHex1);

            A.CustomDash customDash1 = new A.CustomDash();
            A.DashStop dashStop1 = new A.DashStop() {DashLength = 48113, SpaceLength = 119811};
            A.DashStop dashStop2 = new A.DashStop() {DashLength = 48113, SpaceLength = 119811};
            A.DashStop dashStop3 = new A.DashStop() {DashLength = 239623, SpaceLength = 119811};
            A.DashStop dashStop4 = new A.DashStop() {DashLength = 239623, SpaceLength = 119811};
            A.DashStop dashStop5 = new A.DashStop() {DashLength = 239623, SpaceLength = 119811};

            customDash1.Append(dashStop1);
            customDash1.Append(dashStop2);
            customDash1.Append(dashStop3);
            customDash1.Append(dashStop4);
            customDash1.Append(dashStop5);
            A.Round round1 = new A.Round();

            outline1.Append(solidFill1);
            outline1.Append(customDash1);
            outline1.Append(round1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(blipFill1);
            shapeProperties1.Append(outline1);

            Wps.ShapeStyle shapeStyle1 = new Wps.ShapeStyle();
            A.LineReference lineReference1 = new A.LineReference() {Index = (UInt32Value) 0U};
            A.FillReference fillReference1 = new A.FillReference() {Index = (UInt32Value) 0U};
            A.EffectReference effectReference1 = new A.EffectReference() {Index = (UInt32Value) 0U};
            A.FontReference fontReference1 = new A.FontReference() {Index = A.FontCollectionIndexValues.Minor};

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

            TextBoxContent textBoxContent1 = new TextBoxContent();

            Paragraph paragraph2 = new Paragraph();

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() {Val = "FrameContents"};
            OverflowPunctuation overflowPunctuation1 = new OverflowPunctuation() {Val = true};
            BiDi biDi2 = new BiDi() {Val = false};
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines()
                {Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto};
            Indentation indentation1 = new Indentation() {Hanging = "0"};
            Justification justification2 = new Justification() {Val = JustificationValues.Center};

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts2 = new RunFonts() {EastAsia = "Courier New"};

            paragraphMarkRunProperties2.Append(runFonts2);

            paragraphProperties2.Append(paragraphStyleId2);
            paragraphProperties2.Append(overflowPunctuation1);
            paragraphProperties2.Append(biDi2);
            paragraphProperties2.Append(spacingBetweenLines1);
            paragraphProperties2.Append(indentation1);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run();

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts3 = new RunFonts()
            {
                Ascii = "Courier New", HighAnsi = "Courier New", EastAsia = "Courier New", ComplexScript = "Lucida Sans"
            };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Italic italic1 = new Italic() {Val = false};
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript() {Val = false};
            Caps caps1 = new Caps() {Val = false};
            SmallCaps smallCaps1 = new SmallCaps() {Val = false};
            Strike strike1 = new Strike() {Val = false};
            DoubleStrike doubleStrike1 = new DoubleStrike() {Val = false};
            Outline outline2 = new Outline() {Val = false};
            Shadow shadow1 = new Shadow() {Val = false};
            Emboss emboss1 = new Emboss() {Val = false};
            Imprint imprint1 = new Imprint() {Val = false};
            Color color1 = new Color() {Val = "auto"};
            Spacing spacing1 = new Spacing() {Val = 0};
            CharacterScale characterScale1 = new CharacterScale() {Val = 100};
            Kern kern1 = new Kern() {Val = (UInt32Value) 2U};
            Position position1 = new Position() {Val = "0"};
            FontSize fontSize1 = new FontSize() {Val = "36"};
            FontSize fontSize2 = new FontSize() {Val = "36"};
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() {Val = "36"};
            Underline underline1 = new Underline() {Val = UnderlineValues.None};
            VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment()
                {Val = VerticalPositionValues.Baseline};
            Emphasis emphasis1 = new Emphasis() {Val = EmphasisMarkValues.None};
            Languages languages1 = new Languages() {Val = "en-US", EastAsia = "zh-CN", Bidi = "ta-IN"};

            runProperties2.Append(runFonts3);
            runProperties2.Append(bold1);
            runProperties2.Append(boldComplexScript1);
            runProperties2.Append(italic1);
            runProperties2.Append(italicComplexScript1);
            runProperties2.Append(caps1);
            runProperties2.Append(smallCaps1);
            runProperties2.Append(strike1);
            runProperties2.Append(doubleStrike1);
            runProperties2.Append(outline2);
            runProperties2.Append(shadow1);
            runProperties2.Append(emboss1);
            runProperties2.Append(imprint1);
            runProperties2.Append(color1);
            runProperties2.Append(spacing1);
            runProperties2.Append(characterScale1);
            runProperties2.Append(kern1);
            runProperties2.Append(position1);
            runProperties2.Append(fontSize1);
            runProperties2.Append(fontSize2);
            runProperties2.Append(fontSizeComplexScript1);
            runProperties2.Append(underline1);
            runProperties2.Append(verticalTextAlignment1);
            runProperties2.Append(emphasis1);
            runProperties2.Append(languages1);
            Text text1 = new Text();
            text1.Text = "ＡＰＰＲＯＶＥＤ";

            run2.Append(runProperties2);
            run2.Append(text1);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            textBoxContent1.Append(paragraph2);

            textBoxInfo21.Append(textBoxContent1);

            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties()
            {
                LeftInset = 19080, TopInset = 19080, RightInset = 19080, BottomInset = 19080,
                Anchor = A.TextAnchoringTypeValues.Center
            };
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

            textBodyProperties1.Append(noAutoFit1);

            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(shapeStyle1);
            wordprocessingShape1.Append(textBoxInfo21);
            wordprocessingShape1.Append(textBodyProperties1);

            graphicData1.Append(wordprocessingShape1);

            graphic1.Append(graphicData1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(graphic1);

            drawing1.Append(anchor1);

            alternateContentChoice1.Append(drawing1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

            Picture picture1 = new Picture();

            V.Rectangle rectangle1 = new V.Rectangle()
            {
                Id = "shape_0",
                Style =
                    "position:absolute;margin-left:347pt;margin-top:411.1pt;width:151.6pt;height:37.95pt;mso-wrap-style:square;v-text-anchor:middle",
                Stroked = true
            };
            rectangle1.SetAttribute(new OpenXmlAttribute("", "ID", "", "Shape1"));
            rectangle1.SetAttribute(new OpenXmlAttribute("", "path", "",
                "m0,0l-2147483645,0l-2147483645,-2147483646l0,-2147483646xe"));
            V.ImageData imageData1 = new V.ImageData() {DetectMouseClick = true, RelationshipId = "rId200"};
            V.Stroke stroke1 = new V.Stroke()
            {
                Weight = "38160", Color = "black", JoinStyle = V.StrokeJoinStyleValues.Round,
                EndCap = V.StrokeEndCapValues.Flat, DashStyle = "longdashdotdot"
            };

            V.TextBox textBox1 = new V.TextBox();

            TextBoxContent textBoxContent2 = new TextBoxContent();

            Paragraph paragraph3 = new Paragraph();

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() {Val = "FrameContents"};
            OverflowPunctuation overflowPunctuation2 = new OverflowPunctuation() {Val = true};
            BiDi biDi3 = new BiDi() {Val = false};
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines()
                {Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto};
            Indentation indentation2 = new Indentation() {Hanging = "0"};
            Justification justification3 = new Justification() {Val = JustificationValues.Center};

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts4 = new RunFonts() {EastAsia = "Courier New"};

            paragraphMarkRunProperties3.Append(runFonts4);

            paragraphProperties3.Append(paragraphStyleId3);
            paragraphProperties3.Append(overflowPunctuation2);
            paragraphProperties3.Append(biDi3);
            paragraphProperties3.Append(spacingBetweenLines2);
            paragraphProperties3.Append(indentation2);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run3 = new Run();

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts5 = new RunFonts()
            {
                Ascii = "Courier New", HighAnsi = "Courier New", EastAsia = "Courier New", ComplexScript = "Lucida Sans"
            };
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            Italic italic2 = new Italic() {Val = false};
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript() {Val = false};
            Caps caps2 = new Caps() {Val = false};
            SmallCaps smallCaps2 = new SmallCaps() {Val = false};
            Strike strike2 = new Strike() {Val = false};
            DoubleStrike doubleStrike2 = new DoubleStrike() {Val = false};
            Outline outline3 = new Outline() {Val = false};
            Shadow shadow2 = new Shadow() {Val = false};
            Emboss emboss2 = new Emboss() {Val = false};
            Imprint imprint2 = new Imprint() {Val = false};
            Color color2 = new Color() {Val = "auto"};
            Spacing spacing2 = new Spacing() {Val = 0};
            CharacterScale characterScale2 = new CharacterScale() {Val = 100};
            Kern kern2 = new Kern() {Val = (UInt32Value) 2U};
            Position position2 = new Position() {Val = "0"};
            FontSize fontSize3 = new FontSize() {Val = "36"};
            FontSize fontSize4 = new FontSize() {Val = "36"};
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() {Val = "36"};
            Underline underline2 = new Underline() {Val = UnderlineValues.None};
            VerticalTextAlignment verticalTextAlignment2 = new VerticalTextAlignment()
                {Val = VerticalPositionValues.Baseline};
            Emphasis emphasis2 = new Emphasis() {Val = EmphasisMarkValues.None};
            Languages languages2 = new Languages() {Val = "en-US", EastAsia = "zh-CN", Bidi = "ta-IN"};

            runProperties3.Append(runFonts5);
            runProperties3.Append(bold2);
            runProperties3.Append(boldComplexScript2);
            runProperties3.Append(italic2);
            runProperties3.Append(italicComplexScript2);
            runProperties3.Append(caps2);
            runProperties3.Append(smallCaps2);
            runProperties3.Append(strike2);
            runProperties3.Append(doubleStrike2);
            runProperties3.Append(outline3);
            runProperties3.Append(shadow2);
            runProperties3.Append(emboss2);
            runProperties3.Append(imprint2);
            runProperties3.Append(color2);
            runProperties3.Append(spacing2);
            runProperties3.Append(characterScale2);
            runProperties3.Append(kern2);
            runProperties3.Append(position2);
            runProperties3.Append(fontSize3);
            runProperties3.Append(fontSize4);
            runProperties3.Append(fontSizeComplexScript2);
            runProperties3.Append(underline2);
            runProperties3.Append(verticalTextAlignment2);
            runProperties3.Append(emphasis2);
            runProperties3.Append(languages2);
            Text text2 = new Text();
            text2.Text = "ＡＰＰＲＯＶＥＤ";

            run3.Append(runProperties3);
            run3.Append(text2);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            textBoxContent2.Append(paragraph3);

            textBox1.Append(textBoxContent2);
            Wvml.TextWrap textWrap1 = new Wvml.TextWrap() {Type = Wvml.WrapValues.None};

            rectangle1.Append(imageData1);
            rectangle1.Append(stroke1);
            rectangle1.Append(textBox1);
            rectangle1.Append(textWrap1);

            picture1.Append(rectangle1);

            alternateContentFallback1.Append(picture1);

            alternateContent1.Append(alternateContentChoice1);
            alternateContent1.Append(alternateContentFallback1);

            run1.Append(runProperties1);
            run1.Append(alternateContent1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            return paragraph1;
        }
    }
}
