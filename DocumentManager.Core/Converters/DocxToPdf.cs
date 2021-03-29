using DocumentManager.Core.Converters.Handlers;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using System;
using System.IO;

namespace DocumentManager.Core.Converters
{
    public class DocxToPdf
    {
        private readonly DocxToDocx _toDocx;
        private readonly ILogger<DocxToPdf> _logger;

        public DocxToPdf(DocxToDocx toDocx, ILogger<DocxToPdf> logger)
        {
            _toDocx = toDocx;
            _logger = logger;
        }

        internal void Do(string docxSource, string pdfTarget, Placeholders rep)
        {
            try
            {
                var ms = _toDocx.Merge(docxSource, rep);

                var tmpDocxFile = Path.Combine(rep.WorkingLocation, $"{Path.GetFileNameWithoutExtension(pdfTarget)}.docx");

                Extensions.WriteMemoryStreamToDisk(ms, tmpDocxFile);

                // adds watermark if requested
                if (rep.IsWaterMarkNeeded)
                {
                    var options = new WaterMarkOptions {Text = "SAMPLE" };

                    _toDocx.AddWaterMark(tmpDocxFile, tmpDocxFile, options);
                }

                var openOffice = new OpenOfficeHandler(_logger, rep);

                openOffice.Convert(tmpDocxFile, pdfTarget);
            }
            catch (Exception e)
            {
                _logger.LogError(e, "");
            }
            finally
            {
                //Helper.ClearDirectory(rep.WorkingLocation);
            }
        }
    }
}
