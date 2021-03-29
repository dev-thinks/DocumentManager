using DocumentManager.Core;
using DocumentManager.Core.Converters.Handlers;
using DocumentManager.Core.Models;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace DocumentManager.Tests.Pdf
{
    public class PdfMergeTests
    {
        static string executableLocation = Path.GetDirectoryName(
            Assembly.GetExecutingAssembly().Location);

        public static void PerformTest(IServiceProvider services)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            var executor = provider.GetRequiredService<Executor>();

            var table = new TableElement
            {
                TableName = "Invoice",
                RowValues = new Dictionary<string, string[]>() {
                    {"LicenseNumber", new string[] {"AM234", "LSD425874"}},
                    {"LicenseType", new string[] {"Animal <br/> Health", "Lab Technician"}},
                    {"ExpireDate", new string[] {DateTime.Now.AddHours(1).ToShortDateString(), DateTime.Now.AddHours(2).ToShortDateString() } },
                    {"Amount", new string[] {"234", "239"} }
                }
            };

            var placeholders = new Placeholders
            {
                TextPlaceholders = new Dictionary<string, string>
                {
                    { "CustomerName", "Japan" },
                    {"OrgName", "This Org Inc."}
                },
                TablePlaceholders = new List<TableElement> { table }
            };

            var qrImage = Extensions.GetFileAsMemoryStream(Path.Combine(executableLocation, "Pdf\\signature.PNG"));

            var qrImageElement = new ImageElement() { Dpi = 300, MemStream = qrImage };

            placeholders.ImagePlaceholders = new Dictionary<string, ImageElement>
            {
                {"Signature", qrImageElement }
            };

            placeholders.IsWaterMarkNeeded = true;

            executor.Convert("Pdf\\PdfMergeTemplate.docx", "PdfMerge.pdf", placeholders);
        }
    }
}
