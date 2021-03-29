using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentManager.Core.MailMerge;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocumentManager.Core.Converters.Handlers
{
    public partial class DocXHandler
    {
        private readonly MemoryStream _docxMs;
        private readonly Placeholders _rep;
        private readonly ILogger _logger;
        private int _imageCounter;

        public DocXHandler(string docXTemplateFilename, Placeholders placeholders, ILogger logger)
        {
            _docxMs = Extensions.GetFileAsMemoryStream(docXTemplateFilename);
            _rep = placeholders;
            _logger = logger;
        }

        public MemoryStream ReplaceAll()
        {
            if (_rep != null)
            {
                if (_rep.TextPlaceholders?.Count > 0)
                {
                    MergeTextFieldCode();
                }

                if (_rep.HyperlinkPlaceholders?.Count > 0)
                {
                    MergeHyperlinkFieldCode();
                }

                if (_rep.TablePlaceholders?.Count > 0)
                {
                    MergeTableFieldCode();
                }

                if (_rep.ImagePlaceholders?.Count > 0)
                {
                    MergeImageFieldCode();
                }
            }

            _docxMs.Position = 0;

            return _docxMs;
        }

        public MemoryStream MergeTextFieldCode()
        {
            if (_rep.TextPlaceholders == null || _rep.TextPlaceholders.Count == 0)
            {
                _logger.LogWarning("No text placeholder defined here");

                return null;
            }

            using (var doc = WordprocessingDocument.Open(_docxMs, true))
            {
                CleanMarkup(doc);

                foreach (var (key, value) in _rep.TextPlaceholders)
                {
                    doc.GetMergeFields(key).ReplaceWithText(value);
                }
            }

            _docxMs.Position = 0;
            return _docxMs;
        }

        public MemoryStream MergeTableFieldCode()
        {
            if (_rep.TablePlaceholders == null || _rep.TablePlaceholders.Count == 0)
            {
                _logger.LogWarning("No table placeholder defined here");

                return null;
            }

            using (var doc = WordprocessingDocument.Open(_docxMs, true))
            {
                CleanMarkup(doc);

                //Take a Row (one Dictionary) at a time
                foreach (var tableElement in _rep.TablePlaceholders)
                {
                    var trDict = tableElement.RowValues;

                    var trCol0 = trDict.First();

                    var tableRows = doc.MainDocumentPart.Document.Body.Descendants<FieldCode>()
                        .WhereNameIs($"{tableElement.Prefix}:{tableElement.TableName}")
                        .FirstOrDefault()?
                        .Ancestors<TableRow>()
                        .ToList();

                    if (tableRows != null && tableRows.Any())
                    {
                        foreach (var templateRow in tableRows)
                        {
                            var tableFields = templateRow.Descendants<FieldCode>();

                            var textElements = tableFields
                                .Where(t =>
                                    trDict.Keys.Select(key => trCol0.Key).Any(s => t.Text.Contains(s)) &&
                                    t.Ancestors<TableCell>().Any());

                            // Loop through all found rows
                            foreach (var textElement in textElements)
                            {
                                var newTableRows = new List<TableRow>();
                                var tableRow = textElement.Ancestors<TableRow>().First();

                                //Lets create row by row and replace placeholders
                                for (var j = 0; j < trCol0.Value.Length; j++)
                                {
                                    newTableRows.Add((TableRow) tableRow.CloneNode(true));
                                    var tableRowCopy = newTableRows[newTableRows.Count - 1];

                                    var mergeFields = tableRow.Descendants<FieldCode>()
                                        .Where(m => !IsTableRangeField(m, tableElement));

                                    //Cycle through the cells of the row to replace from the Dictionary value ( string array)
                                    foreach (var mergeField in mergeFields)
                                    {
                                        //Now cycle through the "columns" (keys) of the Dictionary and replace item by item
                                        for (var index = 0; index < trDict.Count; index++)
                                        {
                                            var item = trDict.ElementAt(index);

                                            if (mergeField.InnerText.StartsWith(
                                                OpenXmlWordHelpers.GetMergeFieldStartString(item.Key)))
                                            {
                                                // TODO - Need to handle breaks
                                                mergeField.ReplaceWithText(item.Value[j]);
                                            }
                                        }
                                    }

                                    // clearing table range fields, only for this table; not for nested table.
                                    tableRow.Descendants<FieldCode>()
                                        .Where(m => IsTableRangeField(m, tableElement)).ToList()
                                        .ForEach(s => s.ReplaceWithText(""));

                                    tableRow.Parent.InsertAfter(tableRowCopy, tableRow);
                                    tableRow = tableRowCopy;
                                }

                                tableRow.Remove();
                            }
                        }
                    }
                }
            }

            _docxMs.Position = 0;

            return _docxMs;
        }

        private bool IsTableRangeField(FieldCode field, TableElement table)
        {
            return field.InnerText.StartsWith(
                       OpenXmlWordHelpers.GetMergeFieldStartString($"{table.Prefix}:{table.TableName}")) ||
                   field.InnerText.StartsWith(
                       OpenXmlWordHelpers.GetMergeFieldStartString($"{table.Suffix}:{table.TableName}"));
        }

        public MemoryStream MergeHyperlinkFieldCode()
        {
            if (_rep.HyperlinkPlaceholders == null || _rep.HyperlinkPlaceholders.Count == 0)
            {
                return null;
            }

            using (var doc = WordprocessingDocument.Open(_docxMs, true))
            {
                CleanMarkup(doc);

                // Search in body, headers and footers
                var linkMergeFields = doc.MainDocumentPart.Document.Descendants<FieldCode>();

                foreach (var linkMergeField in linkMergeFields)
                {
                    foreach (var replace in _rep.HyperlinkPlaceholders)
                    {
                        var pl =  replace.Key ;
                        if (linkMergeField.Text.Contains(pl))
                        {
                            Run rFldCode = linkMergeField.Parent as Run;
                            Run rBegin = rFldCode?.PreviousSibling<Run>();
                            Run rSep = rFldCode?.NextSibling<Run>();

                            Run rText = rSep?.NextSibling<Run>();
                            Run rEnd = rText?.NextSibling<Run>();

                            rFldCode?.Remove();
                            rBegin?.Remove();
                            rSep?.Remove();
                            rEnd?.Remove();

                            var run = rText;
                            Text t = rText?.GetFirstChild<Text>();

                            if (t != null)
                            {
                                t.Text = string.Empty;
                            }

                            var relation =
                                doc.MainDocumentPart.AddHyperlinkRelationship(
                                    new Uri(replace.Value.Link, UriKind.RelativeOrAbsolute), true);

                            string relationId = relation.Id;
                            var linkText = string.IsNullOrEmpty(replace.Value.Text)
                                ? replace.Value.Link
                                : replace.Value.Text;

                            var hyper =
                                new Hyperlink(
                                    new Run(
                                        new RunProperties(new RunStyle() { Val = "Hyperlink" }),
                                        new Text(linkText)))
                                {
                                    Id = relationId,
                                    History = OnOffValue.FromBoolean(true)
                                };

                            run?.Parent.InsertBefore(hyper, run);
                            // run.Remove();
                        }
                    }
                }
            }

            _docxMs.Position = 0;

            return _docxMs;
        }

        public MemoryStream MergeImageFieldCode()
        {
            using (var doc = WordprocessingDocument.Open(_docxMs, true))
            {
                CleanMarkup(doc);

                var imageMergeFields = doc.MainDocumentPart.Document.Descendants<FieldCode>();

                foreach (var imageMergeField in imageMergeFields)
                {
                    foreach (var replace in _rep.ImagePlaceholders)
                    {
                        string pl =  replace.Key;
                        if (imageMergeField.Text.Contains(pl))
                        {
                            _imageCounter++;

                            Run rFldCode = imageMergeField.Parent as Run;
                            Run rBegin = rFldCode?.PreviousSibling<Run>();
                            Run rSep = rFldCode?.NextSibling<Run>();

                            Run rText = rSep?.NextSibling<Run>();
                            Run rEnd = rText?.NextSibling<Run>();

                            rFldCode?.Remove();
                            rBegin?.Remove();
                            rSep?.Remove();
                            rEnd?.Remove();

                            var run = rText;
                            Text t = rText?.GetFirstChild<Text>();

                            if (t != null)
                            {
                                t.Text = string.Empty;
                            }

                            var imageHandler = new ImageHandler(_logger);
                            imageHandler.AppendImageToElement(replace, run, doc, _imageCounter);
                        }
                    }
                }
            }

            _docxMs.Position = 0;

            return _docxMs;

        }

        private static void CleanMarkup(WordprocessingDocument doc)
        {

        }
    }
}
