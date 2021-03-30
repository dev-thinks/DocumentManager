using DocumentManager.Core;
using DocumentManager.Core.Models;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace DocumentManager.Tests.Watermark
{
    public class WaterMarkTests
    {
        public static void PerformTest(IServiceProvider services)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            var executor = provider.GetRequiredService<Executor>();
            
            var placeholders = new Placeholders
            {
                IsWaterMarkNeeded = true
            };

            executor.Convert("WaterMark\\WaterMarkTemplate.docx", "WithWaterMark.docx", placeholders);

            //executor.RemoveWaterMark("WithWaterMark.docx", "WithOutWaterMark.docx");
        }
    }
}
