using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using SharedXmlToJsonl.Commands;
using SharedXmlToJsonl.Interfaces;

namespace XlsxXmlToJsonl.Commands
{
    /// <summary>
    /// Command handler for converting XLSX files to JSONL format.
    /// </summary>
    public class ConvertXlsxCommand : CommandHandlerBase<ConvertXlsxOptions>
    {
        private readonly IXlsxProcessor _processor;
        private readonly IJsonWriter _jsonWriter;

        /// <summary>
        /// Initializes a new instance of the ConvertXlsxCommand class.
        /// </summary>
        public ConvertXlsxCommand(
            ILogger<ConvertXlsxCommand> logger,
            IServiceProvider serviceProvider,
            IXlsxProcessor processor,
            IJsonWriter jsonWriter)
            : base(logger, serviceProvider)
        {
            _processor = processor ?? throw new ArgumentNullException(nameof(processor));
            _jsonWriter = jsonWriter ?? throw new ArgumentNullException(nameof(jsonWriter));
        }

        /// <summary>
        /// Executes the core conversion logic.
        /// </summary>
        protected override async Task<int> ExecuteCoreAsync(
            ConvertXlsxOptions options,
            CancellationToken cancellationToken)
        {
            _logger.LogInformation("Starting XLSX conversion for {InputPath}", options.InputPath);

            var result = await _processor.ProcessAsync(
                options.InputPath,
                options.OutputDirectory,
                cancellationToken);

            if (result.Success)
            {
                _logger.LogInformation("Successfully processed {ItemsCount} sheets", result.ItemsProcessed);
                return SharedXmlToJsonl.CommonBase.ExitSuccess;
            }
            else
            {
                _logger.LogError("Processing failed: {ErrorMessage}", result.ErrorMessage);
                return SharedXmlToJsonl.CommonBase.ExitProcessingError;
            }
        }

        /// <summary>
        /// Sets up the command with the service provider.
        /// </summary>
        public override void SetupCommand(IServiceProvider serviceProvider)
        {
            // Command setup logic can be added here if needed
        }

        /// <summary>
        /// Pre-processing hook.
        /// </summary>
        protected override Task OnBeforeExecuteAsync(ConvertXlsxOptions options, CancellationToken cancellationToken)
        {
            if (options.Verbose)
            {
                _logger.LogDebug("Options: MaxSheets={MaxSheets}, IncludeHiddenSheets={IncludeHidden}, ExtractFormulas={ExtractFormulas}, ExtractValues={ExtractValues}",
                    options.MaxSheets, options.IncludeHiddenSheets, options.ExtractFormulas, options.ExtractValues);
                _logger.LogDebug("Sheet range: {StartIndex} to {EndIndex}",
                    options.StartSheetIndex, options.EndSheetIndex == -1 ? "all" : options.EndSheetIndex.ToString());
            }
            return Task.CompletedTask;
        }
    }
}
