using DocumentManager.Core.Converters;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.IO;

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

        public void Convert(string source, string target, Placeholders fieldValues = null)
        {
            var exception = ValidateParameterInputFile(source);

            if (exception != null)
            {
                throw exception;
            }

            fieldValues ??= new Placeholders() {OpenOfficeLocation = _config["locationOfLibreOfficeSoffice"]};

            if (source.EndsWith(".docx"))
            {
                if (target.EndsWith(".docx"))
                {
                    _toDocx.Do(source, target, fieldValues);
                }
                else if (target.EndsWith(".pdf"))
                {
                    
                }
                else if (target.EndsWith(".html") || target.EndsWith(".htm"))
                {
                    
                }
            }
            else if (source.EndsWith(".html") || source.EndsWith(".htm"))
            {
                if (target.EndsWith(".html") || target.EndsWith(".htm"))
                {
                    
                }
                else if (target.EndsWith(".docx"))
                {

                    
                }
                else if (target.EndsWith(".pdf"))
                {
                    
                }
            }
        }

        /// <summary>
        /// Adds watermark for the source document.
        /// </summary>
        /// <remarks>If source and target are same, it will replace the source document with watermark</remarks>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="options"></param>
        public void AddWaterMark(string source, string target, WaterMarkOptions options = null)
        {
            var exception = ValidateParameterInputFile(source);

            if (exception != null)
            {
                throw exception;
            }

            options ??= new WaterMarkOptions();

            _toDocx.AddWaterMark(source,target, options);
        }

        /// <summary>
        /// Removes watermark from the source document
        /// </summary>
        /// <remarks>If source and target are same, it will replace the source document without watermark</remarks>
        /// <param name="source"></param>
        /// <param name="target"></param>
        public void RemoveWaterMark(string source, string target)
        {
            var exception = ValidateParameterInputFile(source);

            if (exception != null)
            {
                throw exception;
            }

            _toDocx.RemoveWaterMark(source, target);
        }

        private Exception ValidateParameterInputFile(string inputFile)
        {
            try
            {
                if (string.IsNullOrEmpty(inputFile))
                {
                    _logger.LogError("Input parameter is null or empty: {Parameter}", nameof(inputFile));

                    return new ArgumentNullException(nameof(inputFile));
                }

                if (!new FileInfo(inputFile).Exists)
                {
                    _logger.LogError("Input file not found: {Parameter}", inputFile);

                    return new FileNotFoundException("File Not Found: " + inputFile, inputFile);
                }
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Error while validating input file: {Parameter}", inputFile);

                return e;
            }

            return null;
        }

        private Exception ValidateParameterOutputFile(string outputFile)
        {
            try
            {
                if (string.IsNullOrEmpty(outputFile))
                {
                    _logger.LogError("Input parameter is null or empty: {Parameter}", nameof(outputFile));

                    return new ArgumentNullException(nameof(outputFile));
                }

                var outputFileInfo = new FileInfo(outputFile);
                if (outputFileInfo.Exists)
                {
                    _logger.LogError("File already exists: {Parameter}", outputFile);

                    return new IOException("File already exists : " + outputFile);
                }

                try
                {
                    using (var test = outputFileInfo.OpenWrite())
                    {
                        test.Close();
                    }
                }
                finally
                {
                    outputFileInfo.Delete();
                }
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Error while validating this file: {Parameter}", outputFile);

                return e;
            }

            return null;
        }
    }
}
