// ReSharper disable PossiblyMistakenUseOfParamsMethod
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentManager.Core.MailMerge;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Linq;

namespace DocumentManager.Core.Converters.Handlers
{
    internal class DocxWatermark
    {
        private readonly ILogger _logger;
        private readonly WaterMarkOptions _options;
        private readonly MemoryStream _docxMs;

        private string WaterMarkTypeId => "_x0000_t136";

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
                OpenXmlDocxRef.GenerateHeaderPart1Content(headerPart1, _options, WaterMarkTypeId);
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
                        if (listPic[n - 1].Descendants<Shape>().Count(s => s.Type == $"#{WaterMarkTypeId}") > 0)
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
    }
}
