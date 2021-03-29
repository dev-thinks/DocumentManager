using DocumentManager.Core.Models;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;

namespace DocumentManager.Core.Converters.Handlers
{
    /// <summary>
    /// Open office handler operations
    /// </summary>
    /// <remarks>Credits and base idea taken from here: https://github.com/Reflexe/doc_to_pdf and 
    /// https://stackoverflow.com/questions/30349542/command-libreoffice-headless-convert-to-pdf-test-docx-outdir-pdf-is-not
    /// </remarks>
    public class OpenOfficeHandler
    {
        private readonly ILogger _logger;
        private readonly Placeholders _placeholders;

        public OpenOfficeHandler(ILogger logger, Placeholders placeholders)
        {
            _logger = logger;
            _placeholders = placeholders;
        }

        public void Convert(string inputFile, string outputFile)
        {
            var commandArgs = new List<string>();
            string convertedFile = "";

            var libreOfficePath = _placeholders.OpenOfficeLocation ?? GetLibreOfficePath();

            //Create tmp folder
            var tmpFolder = _placeholders.WorkingLocation; // Path.Combine(_placeholders.WorkingLocation, "OpenOffice");
            if (!Directory.Exists(tmpFolder))
            {
                Directory.CreateDirectory(tmpFolder);
            }

            commandArgs.Add("--convert-to");

            if ((inputFile.EndsWith(".html") || inputFile.EndsWith(".htm")) && outputFile.EndsWith(".pdf"))
            {
                commandArgs.Add("pdf:writer_pdf_Export");
                convertedFile = Path.Combine(tmpFolder, Path.GetFileNameWithoutExtension(inputFile) + ".pdf");
            }
            else if (inputFile.EndsWith(".docx") && outputFile.EndsWith(".pdf"))
            {
                commandArgs.Add("pdf:writer_pdf_Export");
                convertedFile = Path.Combine(tmpFolder, Path.GetFileNameWithoutExtension(inputFile) + ".pdf");
            }
            else if (inputFile.EndsWith(".docx") && (outputFile.EndsWith(".html") || outputFile.EndsWith(".htm")))
            {
                commandArgs.Add("html:HTML:EmbedImages");
                convertedFile = Path.Combine(tmpFolder, Path.GetFileNameWithoutExtension(inputFile) + ".html");
            }
            else if ((inputFile.EndsWith(".html") || inputFile.EndsWith(".htm")) && (outputFile.EndsWith(".docx")))
            {
                commandArgs.Add("docx:\"Office Open XML Text\"");
                convertedFile = Path.Combine(tmpFolder, Path.GetFileNameWithoutExtension(inputFile) + ".docx");
            }

            commandArgs.AddRange(new[] { inputFile, "--norestore", "--writer", "--headless", "--outdir", tmpFolder });

            var procStartInfo = new ProcessStartInfo(libreOfficePath);
            foreach (var arg in commandArgs)
            {
                procStartInfo.ArgumentList.Add(arg);
            }

            procStartInfo.RedirectStandardOutput = true;
            procStartInfo.UseShellExecute = false;
            procStartInfo.CreateNoWindow = true;
            procStartInfo.WorkingDirectory = Environment.CurrentDirectory;

            var process = new Process() { StartInfo = procStartInfo };
            Process[] pname = Process.GetProcessesByName("soffice");

            //Supposedly, only one instance of Libre Office can be run simultaneously
            while (pname.Length > 0)
            {
                Thread.Sleep(5000);
                pname = Process.GetProcessesByName("soffice");
            }

            process.Start();
            process.WaitForExit();

            // Check for failed exit code.
            if (process.ExitCode != 0)
            {
                throw new OpenOfficeHandlerException(process.ExitCode);
            }
            else
            {
                if (File.Exists(outputFile))
                {
                    File.Delete(outputFile);
                }

                if (File.Exists(convertedFile))
                {
                    File.Move(convertedFile, outputFile);
                }

                // Helper.ClearDirectory(tmpFolder);
            }
        }

        private static string GetLibreOfficePath()
        {
            switch (Environment.OSVersion.Platform)
            {
                case PlatformID.Unix:
                    return "/usr/bin/soffice";
                case PlatformID.Win32NT:
                    string binaryDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    return binaryDirectory + "\\Windows\\program\\soffice.exe";
                default:
                    throw new PlatformNotSupportedException("Your OS is not supported");
            }
        }
    }

    public class OpenOfficeHandlerException : Exception
    {
        public OpenOfficeHandlerException(int exitCode)
            : base(string.Format("LibreOffice has failed with " + exitCode))
        { }
    }
}
