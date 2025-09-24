using System;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PptxXmlToJsonl.Commands;
using SharedXmlToJsonl;
using SharedXmlToJsonl.DependencyInjection;

public class Program
{
    public static async Task<int> Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: pptx-xml-to-jsonl <input-file.pptx> <output-directory>");
            return CommonBase.ExitUsageError;
        }

        var inputPath = args[0];
        var outputDirectory = args[1];

        // Create and configure the host
        var host = Host.CreateDefaultBuilder(args)
            .ConfigureServices((context, services) =>
            {
                // Add shared document processing services
                services.AddDocumentProcessing(context.Configuration);

                // Add PPTX-specific services
                services.AddScoped<PptxXmlToJsonl.Processors.PptxProcessor>();
                services.AddScoped<SharedXmlToJsonl.Interfaces.IPptxProcessor, PptxXmlToJsonl.Processors.PptxProcessor>();
                services.AddScoped<ConvertPptxCommand>();
            })
            .Build();

        // Execute the conversion
        using var scope = host.Services.CreateScope();
        var command = scope.ServiceProvider.GetRequiredService<ConvertPptxCommand>();

        var options = new ConvertPptxOptions
        {
            InputPath = inputPath,
            OutputDirectory = outputDirectory,
            Verbose = false
        };

        return await command.ExecuteAsync(options);
    }
}
