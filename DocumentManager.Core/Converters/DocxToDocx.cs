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

        public void Do(string source, string target, Placeholders rep)
        {
            _logger.LogTrace("Start processing docx to docx transformation");

            var ms = Merge(source, rep);

            _logger.LogTrace("Docx to docx Transformation done");

            StreamHandler.WriteMemoryStreamToDisk(ms, target);

            var watermark = new DocxWatermark(_logger, target);
            var ws = watermark.Do();

            StreamHandler.WriteMemoryStreamToDisk(ws, "test.docx");
        }

        public MemoryStream Merge(string source, Placeholders rep)
        {
            var docx = new DocXHandler(source, rep, _logger);
            var ms = docx.ReplaceAll();

            return ms;
        }
    }
}
