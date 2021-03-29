using DocumentManager.Core.Converters.Handlers;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace DocumentManager.Core.Converters
{
    public class DocxToPdf
    {
        private readonly IConfiguration _config;
        private readonly DocxToDocx _toDocx;
        private readonly ILogger<DocxToPdf> _logger;

        public DocxToPdf(IConfiguration config, DocxToDocx toDocx, ILogger<DocxToPdf> logger)
        {
            _config = config;
            _toDocx = toDocx;
            _logger = logger;
        }

        public void Do(string source, string target, Placeholders rep)
        {
            var ms = _toDocx.Merge(source, rep);

            var tmpFile = Path.Combine(Path.GetDirectoryName(target),
                $"{Path.GetFileNameWithoutExtension(target)}{Guid.NewGuid().ToString().Substring(0, 10)}.docx");

            Extensions.WriteMemoryStreamToDisk(ms, tmpFile);

            try
            {
                LibreOfficeWrapper.Convert(tmpFile, target, _config["locationOfLibreOfficeSoffice"]);
            }
            catch (Exception e)
            {
                _logger.LogError(e, "");
            }
            finally
            {

                if (File.Exists(tmpFile))
                {
                    File.Delete(tmpFile);
                }
            }
        }
    }
}
