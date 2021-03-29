using DocumentManager.Core;
using DocumentManager.Core.Converters;

// ReSharper disable once CheckNamespace
namespace Microsoft.Extensions.DependencyInjection
{
    public static class ServiceExtensions
    {
        public static IServiceCollection AddDocumentManager(this IServiceCollection services)
        {
            services.AddScoped<Executor>();
            services.AddScoped<DocxToDocx>();
            services.AddScoped<DocxToPdf>();

            return services;
        }
    }
}
