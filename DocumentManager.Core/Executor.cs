using DocumentManager.Core.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DocumentManager.Core
{
    public class Executor
    {
        private readonly IConfiguration _config;
        private readonly ILogger<Executor> _logger;

        public Executor(IConfiguration config, ILogger<Executor> logger)
        {
            _config = config;
            _logger = logger;
        }

        public void Convert(string inputFile, string outputFile, Placeholders rep = null)
        {
            if (inputFile.EndsWith(".docx"))
            {
                if (outputFile.EndsWith(".docx"))
                {
                    
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
