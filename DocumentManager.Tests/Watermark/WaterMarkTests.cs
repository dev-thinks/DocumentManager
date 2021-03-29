using DocumentManager.Core;
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

            executor.AddWaterMark("WaterMark\\WaterMarkTemplate.docx", "With_WaterMark.docx");

            executor.RemoveWaterMark("With_WaterMark.docx", "Without_WaterMark.docx");
        }
    }
}
