using DocumentManager.Core.Converters;
using DocumentManager.Core.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using DocumentManager.Core.Converters.Handlers;

namespace DocumentManager.Core
{
    public class Executor
    {
        private readonly IConfiguration _config;
        private readonly DocxToDocx _toDocx;
        private readonly DocxToPdf _toPdf;
        private readonly ILogger<Executor> _logger;

        public Executor(IConfiguration config, DocxToDocx toDocx, DocxToPdf toPdf, ILogger<Executor> logger)
        {
            _config = config;
            _toDocx = toDocx;
            _toPdf = toPdf;
            _logger = logger;
        }

        /// <summary>
        /// Converts source template into target document after merging data
        /// </summary>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="fieldValues"></param>
        public void Convert(string source, string target, Placeholders fieldValues = null)
        {
            var exception = ValidateParameterInputFile(source);

            if (exception != null)
            {
                throw exception;
            }

            fieldValues ??= new Placeholders();

            if (source.EndsWith(".docx"))
            {
                if (target.EndsWith(".docx"))
                {
                    _toDocx.Do(source, target, fieldValues);
                }
                else if (target.EndsWith(".pdf"))
                {
                    fieldValues = (Placeholders) AssignDocumentOptions(fieldValues);

                    ValidateAndSetupWorkingDirectory(fieldValues);

                    _toPdf.Do(source, target, fieldValues);
                }
                else if (target.EndsWith(".html") || target.EndsWith(".htm"))
                {
                    fieldValues = (Placeholders) AssignDocumentOptions(fieldValues);

                    ValidateAndSetupWorkingDirectory(fieldValues);
                }
            }
            else if (source.EndsWith(".html") || source.EndsWith(".htm"))
            {
                fieldValues = (Placeholders) AssignDocumentOptions(fieldValues);

                ValidateAndSetupWorkingDirectory(fieldValues);

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
        /// Converts source template into target document after merging data
        /// </summary>
        /// <returns>Returns memory stream of modified document</returns>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="fieldValues"></param>
        /// <returns></returns>
        public MemoryStream Convert(string source, FileType target, Placeholders fieldValues = null)
        {
            var exception = ValidateParameterInputFile(source);

            if (exception != null)
            {
                throw exception;
            }

            fieldValues ??= new Placeholders();

            if (source.EndsWith(".docx"))
            {
                if (target == FileType.Docx)
                {
                    return _toDocx.Merge(source, fieldValues);
                }
                else if (target == FileType.Pdf)
                {
                    fieldValues = (Placeholders) AssignDocumentOptions(fieldValues);
                }
                else if (target == FileType.Html || target == FileType.Htm)
                {
                    fieldValues = (Placeholders) AssignDocumentOptions(fieldValues);
                }
            }
            else if (source.EndsWith(".html") || source.EndsWith(".htm"))
            {
                fieldValues = (Placeholders) AssignDocumentOptions(fieldValues);

                if (target == FileType.Html || target == FileType.Htm)
                {

                }
                else if (target == FileType.Docx)
                {


                }
                else if (target == FileType.Pdf)
                {

                }
            }

            return null;
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

            _toDocx.AddWaterMark(source, target, options);
        }

        /// <summary>
        /// Adds watermark for the source document.
        /// </summary>
        /// <returns>Memory stream of the modified document</returns>
        /// <param name="source"></param>
        /// <param name="options"></param>
        public MemoryStream AddWaterMark(string source, WaterMarkOptions options = null)
        {
            var exception = ValidateParameterInputFile(source);

            if (exception != null)
            {
                throw exception;
            }

            options ??= new WaterMarkOptions();

            var ms = _toDocx.AddWaterMark(source, options);

            return ms;
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

        /// <summary>
        /// Removes watermark from the source document
        /// </summary>
        /// <returns>Returns memory stream of modified document</returns>
        /// <param name="source"></param>
        public MemoryStream RemoveWaterMark(string source)
        {
            var exception = ValidateParameterInputFile(source);

            if (exception != null)
            {
                throw exception;
            }

            var ms = _toDocx.RemoveWaterMark(source);

            return ms;
        }

        /// <summary>
        /// Adds stamp mark for the source document.
        /// </summary>
        /// <remarks>If source and target are same, it will replace the source document with stampmark</remarks>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <param name="options"></param>
        public void AddStampMark(string source, string target, StampMarkOptions options = null)
        {
            var exception = ValidateParameterInputFile(source);

            if (exception != null)
            {
                throw exception;
            }

            options ??= new StampMarkOptions();

            _toDocx.AddStampMark(source, target, options);
        }

        /// <summary>
        /// Merge multiple docx files into single docx
        /// </summary>
        /// <param name="mergedTargetDoc"></param>
        /// <param name="mergeDocs"></param>
        public void MergeDocx(string mergedTargetDoc, params string[] mergeDocs)
        {
            if (mergeDocs == null || mergeDocs.Length == 0)
            {
                throw new Exception("No source document provided to merge.");
            }

            var merger = new DocxMerger(_logger);

            merger.Do(mergedTargetDoc, mergeDocs);
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

        private void ValidateAndSetupWorkingDirectory(DocumentOptions options)
        {
            if (string.IsNullOrEmpty(options.OpenOfficeLocation))
            {
                _logger.LogError("Not a valid open office version: {@Options}", options);

                throw new Exception("Not a valid open office version.");
            }

            if (!File.Exists(options.OpenOfficeLocation))
            {
                _logger.LogError("Open office is not setup correctly: {OfficeLocation}", options.OpenOfficeLocation);

                throw new Exception("Open office is not setup correctly.");
            }

            if (!Directory.Exists(options.WorkingLocation))
            {
                try
                {
                    Directory.CreateDirectory(options.WorkingLocation);
                }
                catch (Exception e)
                {
                    _logger.LogError(e, "Error while setup working directory: {Path}", options.WorkingLocation);

                    throw;
                }
            }
        }

        private DocumentOptions AssignDocumentOptions(DocumentOptions options)
        {
            options.OpenOfficeLocation = _config["OpenOfficeLocation"];
            options.WorkingLocation = $"{_config["WorkingLocation"]}\\{options.CurrentProcessId}\\";

            return options;
        }
    }
}
