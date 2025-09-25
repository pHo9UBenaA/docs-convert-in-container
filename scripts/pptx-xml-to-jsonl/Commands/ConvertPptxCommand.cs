using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using SharedXmlToJsonl.Commands;
using SharedXmlToJsonl.Interfaces;

namespace PptxXmlToJsonl.Commands
{
    /// <summary>
    /// Command handler for converting PPTX files to JSONL format.
    /// </summary>
    public partial class ConvertPptxCommand : CommandHandlerBase<ConvertPptxOptions>
    {
        private readonly IPptxProcessor _processor;
        private readonly IJsonWriter _jsonWriter;

        /// <summary>
        /// Initializes a new instance of the ConvertPptxCommand class.
        /// </summary>
        public ConvertPptxCommand(
            ILogger<ConvertPptxCommand> logger,
            IServiceProvider serviceProvider,
            IPptxProcessor processor,
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
            ConvertPptxOptions options,
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
                LogProcessingError(Logger, result.ErrorMessage ?? "Unknown error");
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

        [LoggerMessage(EventId = 1, Level = LogLevel.Information, Message = "Starting PPTX conversion for {InputPath}")]
        private static partial void LogStartingConversion(ILogger logger, string inputPath);

        [LoggerMessage(EventId = 2, Level = LogLevel.Information, Message = "Successfully processed {ItemsCount} items")]
        private static partial void LogProcessingSuccess(ILogger logger, int itemsCount);

        [LoggerMessage(EventId = 3, Level = LogLevel.Error, Message = "Processing failed: {ErrorMessage}")]
        private static partial void LogProcessingError(ILogger logger, string errorMessage);

        [LoggerMessage(EventId = 4, Level = LogLevel.Debug, Message = "Options: MaxSlides={MaxSlides}, IncludeHiddenSlides={IncludeHidden}, ExtractShapes={ExtractShapes}, ExtractText={ExtractText}")]
        private static partial void LogOptions(ILogger logger, int maxSlides, bool includeHidden, bool extractShapes, bool extractText);

        /// <summary>
        /// Pre-processing hook.
        /// </summary>
        protected override Task OnBeforeExecuteAsync(ConvertPptxOptions options, CancellationToken cancellationToken)
        {
            if (options.Verbose)
            {
                LogOptions(Logger, options.MaxSlides, options.IncludeHiddenSlides, options.ExtractShapes, options.ExtractText);
            }
            return Task.CompletedTask;
        }
    }
}
