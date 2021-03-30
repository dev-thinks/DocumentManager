using DocumentManager.Core;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace DocumentManager.Tests.Merge
{
    public class MergeDocxTests
    {
        public static void PerformTest(IServiceProvider services)
        {
            using IServiceScope serviceScope = services.CreateScope();
            IServiceProvider provider = serviceScope.ServiceProvider;

            var executor = provider.GetRequiredService<Executor>();

            executor.MergeDocx("Merged.docx", new []{ "Merge\\1.docx", "Merge\\2.docx" });
        }
    }
}
