using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentManager.Core.MailMerge;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
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
    }
}
