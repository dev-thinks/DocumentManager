using DocumentManager.Core.Converters.Handlers;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using System.IO;

namespace DocumentManager.Core.Converters
{
    public class DocxToDocx
    {
        private readonly ILogger<DocxToDocx> _logger;

        public DocxToDocx(ILogger<DocxToDocx> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Converts docx template into docx document after replacing data from placeholder model
        /// </summary>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="rep"></param>
        internal void Do(string source, string target, Placeholders rep)
        {
            _logger.LogTrace("Start processing docx to docx transformation");

            var ms = Merge(source, rep);

            _logger.LogTrace("Docx to docx Transformation done");

            Extensions.WriteMemoryStreamToDisk(ms, target);

            if (rep.IsWaterMarkNeeded)
            {
                var options = new WaterMarkOptions
                {
                    Text = rep.WaterMarkText ?? "SAMPLE",
                    WaterMarkFor = FileType.Docx
                };

                AddWaterMark(target, target, options);
            }
        }

        /// <summary>
        /// Merges source docx template with data
        /// </summary>
        /// <remarks>Replaces only the MERGEFIELD field codes in the template</remarks>
        /// <param name="source"></param>
        /// <param name="rep"></param>
        /// <returns></returns>
        internal MemoryStream Merge(string source, Placeholders rep)
        {
            var docx = new DocXHandler(source, rep, _logger);
            var ms = docx.ReplaceAll();

            return ms;
        }

        /// <summary>
        /// Adds watermark for the source document.
        /// </summary>
        /// <remarks>If source and target are same, it will replace the source document with watermark</remarks>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="options"></param>
        internal void AddWaterMark(string source, string target, WaterMarkOptions options)
        {
            var watermark = new DocxWatermark(source, _logger, options);
            var ws = watermark.Do();

            Extensions.WriteMemoryStreamToDisk(ws, target);
        }

        /// <summary>
        /// Adds watermark for the source document
        /// </summary>
        /// <param name="source"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        internal MemoryStream AddWaterMark(string source, WaterMarkOptions options)
        {
            var watermark = new DocxWatermark(source, _logger, options);
            var ws = watermark.Do();

            return ws;
        }

        /// <summary>
        /// Removes watermark from the source document
        /// </summary>
        /// <remarks>If source and target are same, it will replace the source document without watermark</remarks>
        /// <param name="source"></param>
        /// <param name="target"></param>
        internal void RemoveWaterMark(string source, string target)
        {
            var noWaterMark = new DocxWatermark(source, _logger, null);
            var woWaterMark = noWaterMark.Remove();

            Extensions.WriteMemoryStreamToDisk(woWaterMark, target);
        }

        /// <summary>
        /// Removes watermark if present from source document
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        internal MemoryStream RemoveWaterMark(string source)
        {
            var noWaterMark = new DocxWatermark(source, _logger, null);
            var woWaterMark = noWaterMark.Remove();

            return woWaterMark;
        }
    }
}
