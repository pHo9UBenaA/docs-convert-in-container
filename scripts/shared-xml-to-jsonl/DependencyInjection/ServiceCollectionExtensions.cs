using System;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using SharedXmlToJsonl.Commands;
using SharedXmlToJsonl.Configuration;
using SharedXmlToJsonl.ErrorHandling;
using SharedXmlToJsonl.Factories;
using SharedXmlToJsonl.Interfaces;
using SharedXmlToJsonl.Processing;
using SharedXmlToJsonl.Repositories;
using SharedXmlToJsonl.Resources;
using SharedXmlToJsonl.Services;

namespace SharedXmlToJsonl.DependencyInjection;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddDocumentProcessing(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(services);
        ArgumentNullException.ThrowIfNull(configuration);

        // Configuration
        services.Configure<ProcessingOptions>(
            configuration.GetSection("Processing"));

        // Factories
        services.AddSingleton<IDocumentProcessorFactory, DocumentProcessorFactory>();
        services.AddSingleton<IElementFactory, ElementFactory>();

        // Services
        services.AddScoped<IJsonWriter, JsonWriter>();
        services.AddScoped<IPackageReader, PackageReader>();
        services.AddScoped<IXmlParser, XmlParser>();

        // Repositories
        services.AddScoped<IDocumentRepository, FileSystemDocumentRepository>();

        // Processing
        services.AddScoped<IParallelProcessor, ParallelProcessor>();
        services.AddScoped<IBufferManager, BufferManager>();

        // Error Handling
        services.AddSingleton<IGlobalErrorHandler, GlobalErrorHandler>();

        // Resources
        services.AddSingleton<IResourceService, ResourceService>();

        // Logging
        services.AddLogging(builder =>
        {
            builder.AddConsole();
            builder.SetMinimumLevel(LogLevel.Information);
        });

        return services;
    }

    public static IServiceCollection AddPptxProcessing(
        this IServiceCollection services)
    {
        ArgumentNullException.ThrowIfNull(services);

        // Register PPTX-specific processors and commands
        // These will be registered in the PPTX project's Program.cs
        return services;
    }

    public static IServiceCollection AddXlsxProcessing(
        this IServiceCollection services)
    {
        ArgumentNullException.ThrowIfNull(services);

        // Register XLSX-specific processors and commands
        // These will be registered in the XLSX project's Program.cs
        return services;
    }
}
