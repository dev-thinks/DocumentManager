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
        public void Do(string source, string target, Placeholders rep)
        {
            _logger.LogTrace("Start processing docx to docx transformation");

            var ms = Merge(source, rep);

            _logger.LogTrace("Docx to docx Transformation done");

            StreamHandler.WriteMemoryStreamToDisk(ms, target);
        }

        /// <summary>
        /// Merges source docx template with data
        /// <remarks>Replaces only the MERGEFIELD field codes in the template</remarks>
        /// </summary>
        /// <param name="source"></param>
        /// <param name="rep"></param>
        /// <returns></returns>
        public MemoryStream Merge(string source, Placeholders rep)
        {
            var docx = new DocXHandler(source, rep, _logger);
            var ms = docx.ReplaceAll();

            return ms;
        }

        /// <summary>
        /// Adds watermark for the source document.
        /// <remarks>If source and target are same, it will replace the source document with watermark</remarks>
        /// </summary>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="options"></param>
        public void AddWaterMark(string source, string target, WaterMarkOptions options)
        {
            var watermark = new DocxWatermark(source, _logger);
            var ws = watermark.Do();

            StreamHandler.WriteMemoryStreamToDisk(ws, target);
        }

        /// <summary>
        /// Removes watermark from the source document
        /// <remarks>If source and target are same, it will replace the source document without watermark</remarks>
        /// </summary>
        /// <param name="source"></param>
        /// <param name="target"></param>
        public void RemoveWaterMark(string source, string target)
        {
            var noWaterMark = new DocxWatermark(source, _logger);
            var woWaterMark = noWaterMark.Remove();

            StreamHandler.WriteMemoryStreamToDisk(woWaterMark, target);
        }
    }
}
