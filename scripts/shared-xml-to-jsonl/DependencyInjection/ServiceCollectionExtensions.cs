using System;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using SharedXmlToJsonl.Commands;
using SharedXmlToJsonl.Configuration;
using SharedXmlToJsonl.Factories;
using SharedXmlToJsonl.Interfaces;
using SharedXmlToJsonl.Services;

namespace SharedXmlToJsonl.DependencyInjection;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddDocumentProcessing(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));

        if (configuration == null)
            throw new ArgumentNullException(nameof(configuration));

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

        // Processors need to be registered in the specific project

        return services;
    }

    // PPTX and XLSX specific registrations should be done in their respective projects
}