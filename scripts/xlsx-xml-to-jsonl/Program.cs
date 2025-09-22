using System;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using XlsxXmlToJsonl.Commands;
using SharedXmlToJsonl;
using SharedXmlToJsonl.DependencyInjection;

public class Program
{
    public static async Task<int> Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: xlsx-xml-to-jsonl <input-file.xlsx> <output-directory>");
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

                // Add XLSX-specific services
                services.AddScoped<XlsxXmlToJsonl.Processors.XlsxProcessor>();
                services.AddScoped<SharedXmlToJsonl.Interfaces.IXlsxProcessor, XlsxXmlToJsonl.Processors.XlsxProcessor>();
                services.AddScoped<ConvertXlsxCommand>();
            })
            .Build();

        // Execute the conversion
        using var scope = host.Services.CreateScope();
        var command = scope.ServiceProvider.GetRequiredService<ConvertXlsxCommand>();

        var options = new ConvertXlsxOptions
        {
            InputPath = inputPath,
            OutputDirectory = outputDirectory,
            Verbose = false
        };

        return await command.ExecuteAsync(options);
    }
}