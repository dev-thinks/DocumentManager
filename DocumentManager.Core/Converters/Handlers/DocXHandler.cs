﻿using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentManager.Core.MailMerge;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;

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
            _docxMs = StreamHandler.GetFileAsMemoryStream(docXTemplateFilename);
            _rep = placeholders;
            _logger = logger;
        }

        public MemoryStream ReplaceAll()
        {
            if (_rep != null)
            {
                MergeTextFieldCode();

                //ReplaceHyperlinks();

                MergeTableFieldCode();

                //ReplaceImages();
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

        public MemoryStream MergeHyperlinks()
        {
            if (_rep.HyperlinkPlaceholders == null || _rep.HyperlinkPlaceholders.Count == 0)
            {
                return null;
            }

            using (var doc = WordprocessingDocument.Open(_docxMs, true))
            {
                CleanMarkup(doc);

                // Search in body, headers and footers
                var documentTexts = doc.MainDocumentPart.Document.Descendants<Text>();

                foreach (var text in documentTexts)
                {
                    foreach (var replace in _rep.HyperlinkPlaceholders)
                    {
                        var pl = _rep.HyperlinkPlaceholderStartTag + replace.Key + _rep.HyperlinkPlaceholderEndTag;
                        if (text.Text.Contains(pl))
                        {
                            var run = text.Ancestors<Run>().First();

                            if (text.Text.StartsWith(pl))
                            {
                                var newAfterRun = (Run)run.Clone();
                                string afterText = text.Text.Substring(pl.Length, text.Text.Length - pl.Length);
                                Text newAfterRunText = newAfterRun.GetFirstChild<Text>();
                                newAfterRunText.Space = SpaceProcessingModeValues.Preserve;
                                newAfterRunText.Text = afterText;

                                run.Parent.InsertAfter(newAfterRun, run);
                            }
                            else if (text.Text.EndsWith(pl))
                            {
                                var newBeforeRun = (Run)run.Clone();
                                string beforeText = text.Text.Substring(0, text.Text.Length - pl.Length);
                                Text newBeforeRunText = newBeforeRun.GetFirstChild<Text>();
                                newBeforeRunText.Space = SpaceProcessingModeValues.Preserve;
                                newBeforeRunText.Text = beforeText;

                                run.Parent.InsertBefore(newBeforeRun, run);
                            }
                            else
                            {
                                //Break the texts into the part before and after image. Then create separate runs for them
                                var pos = text.Text.IndexOf(pl, StringComparison.CurrentCulture);

                                var newBeforeRun = (Run)run.Clone();
                                string beforeText = text.Text.Substring(0, pos);
                                Text newBeforeRunText = newBeforeRun.GetFirstChild<Text>();
                                newBeforeRunText.Space = SpaceProcessingModeValues.Preserve;
                                newBeforeRunText.Text = beforeText;
                                run.Parent.InsertBefore(newBeforeRun, run);

                                var newAfterRun = (Run)run.Clone();
                                string afterText =
                                    text.Text.Substring(pos + pl.Length, text.Text.Length - pos - pl.Length);
                                Text newAfterRunText = newAfterRun.GetFirstChild<Text>();
                                newAfterRunText.Space = SpaceProcessingModeValues.Preserve;
                                newAfterRunText.Text = afterText;
                                run.Parent.InsertAfter(newAfterRun, run);
                            }

                            var relation =
                                doc.MainDocumentPart.AddHyperlinkRelationship(
                                    new Uri(replace.Value.Link, UriKind.RelativeOrAbsolute), true);
                            string relationid = relation.Id;
                            var linkText = string.IsNullOrEmpty(replace.Value.Text)
                                ? replace.Value.Link
                                : replace.Value.Text;
                            var hyper =
                                new Hyperlink(
                                    new Run(
                                        new RunProperties(new RunStyle() { Val = "Hyperlink" }),
                                        new Text(linkText)))
                                {
                                    Id = relationid,
                                    History = OnOffValue.FromBoolean(true)
                                };

                            run.Parent.InsertBefore(hyper, run);
                            run.Remove();
                        }
                    }
                }
            }

            _docxMs.Position = 0;
            return _docxMs;
        }
    }
}
