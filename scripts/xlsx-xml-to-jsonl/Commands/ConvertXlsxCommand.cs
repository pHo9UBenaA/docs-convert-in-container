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
    public partial class ConvertXlsxCommand : CommandHandlerBase<ConvertXlsxOptions>
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
            LogStartingConversion(Logger, options.InputPath);

            var result = await _processor.ProcessAsync(
                options.InputPath,
                options.OutputDirectory,
                cancellationToken);

            if (result.Success)
            {
                LogProcessingSuccess(Logger, result.ItemsProcessed);
                return SharedXmlToJsonl.CommonBase.ExitSuccess;
            }
            else
            {
                LogProcessingError(Logger, result.ErrorMessage);
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

        [LoggerMessage(EventId = 1, Level = LogLevel.Information, Message = "Starting XLSX conversion for {InputPath}")]
        private static partial void LogStartingConversion(ILogger logger, string inputPath);

        [LoggerMessage(EventId = 2, Level = LogLevel.Information, Message = "Successfully processed {ItemsCount} sheets")]
        private static partial void LogProcessingSuccess(ILogger logger, int itemsCount);

        [LoggerMessage(EventId = 3, Level = LogLevel.Error, Message = "Processing failed: {ErrorMessage}")]
        private static partial void LogProcessingError(ILogger logger, string errorMessage);

        [LoggerMessage(EventId = 4, Level = LogLevel.Debug, Message = "Options: MaxSheets={MaxSheets}, IncludeHiddenSheets={IncludeHidden}, ExtractFormulas={ExtractFormulas}, ExtractValues={ExtractValues}")]
        private static partial void LogOptionsMain(ILogger logger, int maxSheets, bool includeHidden, bool extractFormulas, bool extractValues);

        [LoggerMessage(EventId = 5, Level = LogLevel.Debug, Message = "Sheet range: {StartIndex} to {EndIndex}")]
        private static partial void LogSheetRange(ILogger logger, int startIndex, string endIndex);

        /// <summary>
        /// Pre-processing hook.
        /// </summary>
        protected override Task OnBeforeExecuteAsync(ConvertXlsxOptions options, CancellationToken cancellationToken)
        {
            if (options.Verbose)
            {
                LogOptionsMain(Logger, options.MaxSheets, options.IncludeHiddenSheets, options.ExtractFormulas, options.ExtractValues);
                LogSheetRange(Logger, options.StartSheetIndex, options.EndSheetIndex == -1 ? "all" : options.EndSheetIndex.ToString());
            }
            return Task.CompletedTask;
        }
    }
}
