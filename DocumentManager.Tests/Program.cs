using DocumentManager.Core;
using DocumentManager.Core.Converters;
using DocumentManager.Tests.HyperLinks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using DocumentManager.Tests.Image;
using DocumentManager.Tests.Watermark;

namespace DocumentManager.Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            using IHost host = CreateHostBuilder(args).Build();

            Console.WriteLine("Hello World!");

            // TableMerge.PerformTest(host.Services);

            // ImageMerge.PerformTest(host.Services);

            DocumentWaterMark.PerformTest(host.Services);

            Console.WriteLine("Completed.");
        }

        private static IHostBuilder CreateHostBuilder(string[] args)
        {
            var hb = Host.CreateDefaultBuilder(args)
                .ConfigureLogging(builder => builder.AddConsole());

            hb.ConfigureServices((_, services) =>
            {
                services.AddScoped<Executor>();
                services.AddScoped<DocxToDocx>();
            });

            return hb;
        }
    }
}
