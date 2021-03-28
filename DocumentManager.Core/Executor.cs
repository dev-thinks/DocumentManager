using DocumentManager.Core.Converters;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DocumentManager.Core
{
    public class Executor
    {
        private readonly IConfiguration _config;
        private readonly DocxToDocx _toDocx;
        private readonly ILogger<Executor> _logger;

        public Executor(IConfiguration config, DocxToDocx toDocx, ILogger<Executor> logger)
        {
            _config = config;
            _toDocx = toDocx;
            _logger = logger;
        }

        public void Convert(string inputFile, string outputFile, Placeholders rep = null)
        {
            if (inputFile.EndsWith(".docx"))
            {
                if (outputFile.EndsWith(".docx"))
                {
                    _toDocx.Do(inputFile, outputFile, rep);
                }
                else if (outputFile.EndsWith(".pdf"))
                {
                    
                }
                else if (outputFile.EndsWith(".html") || outputFile.EndsWith(".htm"))
                {
                    
                }
            }
            else if (inputFile.EndsWith(".html") || inputFile.EndsWith(".htm"))
            {
                if (outputFile.EndsWith(".html") || outputFile.EndsWith(".htm"))
                {
                    
                }
                else if (outputFile.EndsWith(".docx"))
                {

                    
                }
                else if (outputFile.EndsWith(".pdf"))
                {
                    
                }
            }
        }
    }
}
