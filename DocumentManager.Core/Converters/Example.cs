using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Op = DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;

namespace GeneratedCode
{
    public class GeneratedClass
    {
        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId2");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            HeaderPart headerPart1 = mainDocumentPart1.AddNewPart<HeaderPart>("rId2");
            GenerateHeaderPart1Content(headerPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId3");
            GenerateFontTablePart1Content(fontTablePart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId4");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "0";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "LibreOffice/7.1.1.2$Windows_x86 LibreOffice_project/fe0b08f4af1bacafe4c7ecc87ce55bb426164676";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "15.0000";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "0";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "0";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "0";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "0";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(application1);
            properties1.Append(applicationVersion1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(paragraphs1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Normal" };
            BiDi biDi1 = new BiDi() { Val = false };
            Justification justification1 = new Justification() { Val = JustificationValues.Left };
            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(biDi1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();
            RunProperties runProperties1 = new RunProperties();

            run1.Append(runProperties1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            SectionProperties sectionProperties1 = new SectionProperties();
            HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "rId2" };
            SectionType sectionType1 = new SectionType() { Val = SectionMarkValues.NextPage };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1417, Right = (UInt32Value)1134U, Bottom = 1134, Left = (UInt32Value)1134U, Header = (UInt32Value)1134U, Footer = (UInt32Value)0U, Gutter = (UInt32Value)0U };
            PageNumberType pageNumberType1 = new PageNumberType() { Format = NumberFormatValues.Decimal };
            FormProtection formProtection1 = new FormProtection() { Val = false };
            TextDirection textDirection1 = new TextDirection() { Val = TextDirectionValues.LefToRightTopToBottom };

            sectionProperties1.Append(headerReference1);
            sectionProperties1.Append(sectionType1);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(pageNumberType1);
            sectionProperties1.Append(formProtection1);
            sectionProperties1.Append(textDirection1);

            body1.Append(paragraph1);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "NSimSun", ComplexScript = "Lucida Sans" };
            Kern kern1 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "zh-CN", Bidi = "hi-IN" };

            runPropertiesBaseStyle1.Append(runFonts1);
            runPropertiesBaseStyle1.Append(kern1);
            runPropertiesBaseStyle1.Append(fontSize1);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript1);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            WidowControl widowControl1 = new WidowControl();
            SuppressAutoHyphens suppressAutoHyphens1 = new SuppressAutoHyphens() { Val = true };

            paragraphPropertiesBaseStyle1.Append(widowControl1);
            paragraphPropertiesBaseStyle1.Append(suppressAutoHyphens1);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal" };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl2 = new WidowControl();
            BiDi biDi2 = new BiDi() { Val = false };

            styleParagraphProperties1.Append(widowControl2);
            styleParagraphProperties1.Append(biDi2);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "NSimSun", ComplexScript = "Lucida Sans" };
            Color color1 = new Color() { Val = "auto" };
            Kern kern2 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };
            Languages languages2 = new Languages() { Val = "en-US", EastAsia = "zh-CN", Bidi = "hi-IN" };

            styleRunProperties1.Append(runFonts2);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(kern2);
            styleRunProperties1.Append(fontSize2);
            styleRunProperties1.Append(fontSizeComplexScript2);
            styleRunProperties1.Append(languages2);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading" };
            StyleName styleName2 = new StyleName() { Val = "Heading" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext() { Val = true };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "120" };

            styleParagraphProperties2.Append(keepNext1);
            styleParagraphProperties2.Append(spacingBetweenLines1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Microsoft YaHei", ComplexScript = "Lucida Sans" };
            FontSize fontSize3 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties2.Append(runFonts3);
            styleRunProperties2.Append(fontSize3);
            styleRunProperties2.Append(fontSizeComplexScript3);

            style2.Append(styleName2);
            style2.Append(basedOn1);
            style2.Append(nextParagraphStyle1);
            style2.Append(primaryStyle2);
            style2.Append(styleParagraphProperties2);
            style2.Append(styleRunProperties2);

            Style style3 = new Style() { Type = StyleValues.Paragraph, StyleId = "TextBody" };
            StyleName styleName3 = new StyleName() { Val = "Body Text" };
            BasedOn basedOn2 = new BasedOn() { Val = "Normal" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Before = "0", After = "140", Line = "276", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties3.Append(spacingBetweenLines2);
            StyleRunProperties styleRunProperties3 = new StyleRunProperties();

            style3.Append(styleName3);
            style3.Append(basedOn2);
            style3.Append(styleParagraphProperties3);
            style3.Append(styleRunProperties3);

            Style style4 = new Style() { Type = StyleValues.Paragraph, StyleId = "List" };
            StyleName styleName4 = new StyleName() { Val = "List" };
            BasedOn basedOn3 = new BasedOn() { Val = "TextBody" };
            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts4 = new RunFonts() { ComplexScript = "Lucida Sans" };

            styleRunProperties4.Append(runFonts4);

            style4.Append(styleName4);
            style4.Append(basedOn3);
            style4.Append(styleParagraphProperties4);
            style4.Append(styleRunProperties4);

            Style style5 = new Style() { Type = StyleValues.Paragraph, StyleId = "Caption" };
            StyleName styleName5 = new StyleName() { Val = "Caption" };
            BasedOn basedOn4 = new BasedOn() { Val = "Normal" };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers1 = new SuppressLineNumbers();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Before = "120", After = "120" };

            styleParagraphProperties5.Append(suppressLineNumbers1);
            styleParagraphProperties5.Append(spacingBetweenLines3);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts5 = new RunFonts() { ComplexScript = "Lucida Sans" };
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            FontSize fontSize4 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties5.Append(runFonts5);
            styleRunProperties5.Append(italic1);
            styleRunProperties5.Append(italicComplexScript1);
            styleRunProperties5.Append(fontSize4);
            styleRunProperties5.Append(fontSizeComplexScript4);

            style5.Append(styleName5);
            style5.Append(basedOn4);
            style5.Append(primaryStyle3);
            style5.Append(styleParagraphProperties5);
            style5.Append(styleRunProperties5);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "Index" };
            StyleName styleName6 = new StyleName() { Val = "Index" };
            BasedOn basedOn5 = new BasedOn() { Val = "Normal" };
            PrimaryStyle primaryStyle4 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers2 = new SuppressLineNumbers();

            styleParagraphProperties6.Append(suppressLineNumbers2);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts6 = new RunFonts() { ComplexScript = "Lucida Sans" };

            styleRunProperties6.Append(runFonts6);

            style6.Append(styleName6);
            style6.Append(basedOn5);
            style6.Append(primaryStyle4);
            style6.Append(styleParagraphProperties6);
            style6.Append(styleRunProperties6);

            Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "HeaderandFooter" };
            StyleName styleName7 = new StyleName() { Val = "Header and Footer" };
            BasedOn basedOn6 = new BasedOn() { Val = "Normal" };
            PrimaryStyle primaryStyle5 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties7 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers3 = new SuppressLineNumbers();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Clear, Position = 709 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Center, Leader = TabStopLeaderCharValues.None, Position = 4986 };
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.None, Position = 9972 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            tabs1.Append(tabStop3);

            styleParagraphProperties7.Append(suppressLineNumbers3);
            styleParagraphProperties7.Append(tabs1);
            StyleRunProperties styleRunProperties7 = new StyleRunProperties();

            style7.Append(styleName7);
            style7.Append(basedOn6);
            style7.Append(primaryStyle5);
            style7.Append(styleParagraphProperties7);
            style7.Append(styleRunProperties7);

            Style style8 = new Style() { Type = StyleValues.Paragraph, StyleId = "Header" };
            StyleName styleName8 = new StyleName() { Val = "Header" };
            BasedOn basedOn7 = new BasedOn() { Val = "HeaderandFooter" };

            StyleParagraphProperties styleParagraphProperties8 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers4 = new SuppressLineNumbers();

            styleParagraphProperties8.Append(suppressLineNumbers4);
            StyleRunProperties styleRunProperties8 = new StyleRunProperties();

            style8.Append(styleName8);
            style8.Append(basedOn7);
            style8.Append(styleParagraphProperties8);
            style8.Append(styleRunProperties8);

            styles1.Append(docDefaults1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of headerPart1.
        private void GenerateHeaderPart1Content(HeaderPart headerPart1)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            Paragraph paragraph2 = new Paragraph();

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Header" };
            SuppressLineNumbers suppressLineNumbers5 = new SuppressLineNumbers();
            BiDi biDi3 = new BiDi() { Val = false };
            Justification justification2 = new Justification() { Val = JustificationValues.Left };
            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();

            paragraphProperties2.Append(paragraphStyleId2);
            paragraphProperties2.Append(suppressLineNumbers5);
            paragraphProperties2.Append(biDi3);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run();
            RunProperties runProperties2 = new RunProperties();

            Picture picture1 = new Picture();

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "shapetype_136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@9,l@10,em@11,21600l@12,21600e" };
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula() { Equation = "val #0" };
            V.Formula formula2 = new V.Formula() { Equation = "sum @0 0 10800" };
            V.Formula formula3 = new V.Formula() { Equation = "sum @0 0 0" };
            V.Formula formula4 = new V.Formula() { Equation = "sum width 0 @0" };
            V.Formula formula5 = new V.Formula() { Equation = "prod @2 2 1" };
            V.Formula formula6 = new V.Formula() { Equation = "prod @3 2 1" };
            V.Formula formula7 = new V.Formula() { Equation = "if @1 @5 @4" };
            V.Formula formula8 = new V.Formula() { Equation = "sum 0 @6 0" };
            V.Formula formula9 = new V.Formula() { Equation = "sum width 0 @6" };
            V.Formula formula10 = new V.Formula() { Equation = "if @1 0 @8" };
            V.Formula formula11 = new V.Formula() { Equation = "if @1 @7 width" };
            V.Formula formula12 = new V.Formula() { Equation = "if @1 @8 0" };
            V.Formula formula13 = new V.Formula() { Equation = "if @1 width @7" };

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
            V.ShapeHandle shapeHandle1 = new V.ShapeHandle() { Position = "@0,21600" };

            shapeHandles1.Append(shapeHandle1);

            shapetype1.Append(stroke1);
            shapetype1.Append(formulas1);
            shapetype1.Append(shapeHandles1);

            V.Shape shape1 = new V.Shape() { Id = "PowerPlusWaterMarkObject", Style = "position:absolute;margin-left:0.05pt;margin-top:256.9pt;width:498.5pt;height:150.4pt;mso-wrap-style:none;v-text-anchor:middle;rotation:315;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "shape_0", FillColor = "silver", Stroked = false, Type = "shapetype_136" };
            V.Path path1 = new V.Path() { AllowTextPath = true };
            V.TextPath textPath1 = new V.TextPath() { Style = "font-family:\"Calibri\";font-size:1pt", On = true, FitShape = true, Trim = true, String = "SAMPLE" };
            V.Fill fill1 = new V.Fill() { Type = V.FillTypeValues.Solid, Opacity = "0.5", Color2 = "#3f3f3f", DetectMouseClick = true };
            V.Stroke stroke2 = new V.Stroke() { Color = "#3465a4", JoinStyle = V.StrokeJoinStyleValues.Round, EndCap = V.StrokeEndCapValues.Flat };
            Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { Type = Wvml.WrapValues.None };

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

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts();
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            Font font1 = new Font() { Name = "Times New Roman" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };

            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);

            Font font2 = new Font() { Name = "Symbol" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "02" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };

            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);

            Font font3 = new Font() { Name = "Arial" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };

            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);

            fontTablePart1.Fonts = fonts1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings();
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "100" };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 709 };
            AutoHyphenation autoHyphenation1 = new AutoHyphenation() { Val = true };

            Compatibility compatibility1 = new Compatibility();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" };

            compatibility1.Append(compatibilitySetting1);

            settings1.Append(zoom1);
            settings1.Append(defaultTabStop1);
            settings1.Append(autoHyphenation1);
            settings1.Append(compatibility1);

            documentSettingsPart1.Settings = settings1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "";
            document.PackageProperties.Title = "";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.Revision = "1";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2021-03-29T18:16:36Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2021-03-29T18:17:27Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "";
            document.PackageProperties.Language = "en-US";
        }



    }
}
