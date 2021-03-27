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
            var ms = Merge(source, rep);

            StreamHandler.WriteMemoryStreamToDisk(ms, target);
        }

        public MemoryStream Merge(string source, Placeholders rep)
        {
            var docx = new DocXHandler(source, rep, _logger);
            var ms = docx.ReplaceAll();

            return ms;
        }
    }
}
