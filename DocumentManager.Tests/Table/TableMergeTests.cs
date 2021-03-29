using DocumentManager.Core.Converters;
using DocumentManager.Core.Models;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;

namespace DocumentManager.Tests.Table
{
    public class TableMergeTests
    {
        public static void PerformTest(IServiceProvider services)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            var docxToDocx = provider.GetRequiredService<DocxToDocx>();

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

            docxToDocx.Do("Table\\InvoiceTable.docx", "InvoiceTableDocument.docx", placeholders);
        }
    }
}
