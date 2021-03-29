using DocumentManager.Core;
using DocumentManager.Core.Converters.Handlers;
using DocumentManager.Core.Models;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace DocumentManager.Tests.Image
{
    public class ImageMergeTests
    {
       static string executableLocation = Path.GetDirectoryName(
            Assembly.GetExecutingAssembly().Location);

        public static void PerformTest(IServiceProvider services)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            var executor = provider.GetRequiredService<Executor>();

            var placeholders = new Placeholders
            {
                TextPlaceholders = new Dictionary<string, string>
                {
                    { "CustomerName", "Japan" },
                    {"OrgName", "This Org Inc."},
                    {"CartCount", "3" }
                },
                HyperlinkPlaceholders = new Dictionary<string, HyperlinkElement>
                {
                    { "PortalUrl", new HyperlinkElement { Text = "PortalUrl", Link = "https://www.voltron.com/"} }
                }
            };

            var qrImage = Extensions.GetFileAsMemoryStream(Path.Combine(executableLocation, "Image\\QRCode.PNG"));

            var qrImageElement = new ImageElement() { Dpi = 300, MemStream = qrImage };

            placeholders.ImagePlaceholders = new Dictionary<string, ImageElement>
            {
                {"Signature", qrImageElement }
            };

            executor.Convert("Image\\ImageMergeTemplate.docx", "ImageMergeDocument.docx", placeholders);
        }
    }
}
