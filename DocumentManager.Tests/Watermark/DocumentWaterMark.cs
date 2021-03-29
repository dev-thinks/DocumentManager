﻿using DocumentManager.Core.Converters;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace DocumentManager.Tests.Watermark
{
    public class DocumentWaterMark
    {
        public static void PerformTest(IServiceProvider services)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            var docxToDocx = provider.GetRequiredService<DocxToDocx>();

            docxToDocx.AddWaterMark("WaterMark\\WaterMark.docx", "With_WaterMark.docx");

            docxToDocx.RemoveWaterMark("With_WaterMark.docx", "Without_WaterMark.docx");
        }
    }
}
