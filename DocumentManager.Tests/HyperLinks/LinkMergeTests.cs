using DocumentManager.Core;
using DocumentManager.Core.Models;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;

namespace DocumentManager.Tests.HyperLinks
{
    public class LinkMergeTests
    {
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

            executor.Convert("HyperLinks\\LinkMergeTemplate.docx", "LinkMergeDocument.docx", placeholders);
        }
    }
}
