using System;
using System.Collections.Generic;
using DocumentManager.Core.Converters;
using DocumentManager.Core.Models;
using Microsoft.Extensions.DependencyInjection;

namespace DocumentManager.Tests.HyperLinks
{
    public class LinkMerge
    {
        public static void PerformTest(IServiceProvider services)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            var docxToDocx = provider.GetRequiredService<DocxToDocx>();

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

            docxToDocx.Do("HyperLinks\\CartReport.docx", "CartReport_Merged.docx", placeholders);
        }
    }
}
