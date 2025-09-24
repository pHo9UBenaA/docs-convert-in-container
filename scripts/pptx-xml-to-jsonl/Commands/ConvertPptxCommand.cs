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
    public class ConvertPptxCommand : CommandHandlerBase<ConvertPptxOptions>
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
            Logger.LogInformation("Starting PPTX conversion for {InputPath}", options.InputPath);

            var result = await _processor.ProcessAsync(
                options.InputPath,
                options.OutputDirectory,
                cancellationToken);

            if (result.Success)
            {
                Logger.LogInformation("Successfully processed {ItemsCount} items", result.ItemsProcessed);
                return SharedXmlToJsonl.CommonBase.ExitSuccess;
            }
            else
            {
                Logger.LogError("Processing failed: {ErrorMessage}", result.ErrorMessage);
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
        protected override Task OnBeforeExecuteAsync(ConvertPptxOptions options, CancellationToken cancellationToken)
        {
            if (options.Verbose)
            {
                Logger.LogDebug("Options: MaxSlides={MaxSlides}, IncludeHiddenSlides={IncludeHidden}, ExtractShapes={ExtractShapes}, ExtractText={ExtractText}",
                    options.MaxSlides, options.IncludeHiddenSlides, options.ExtractShapes, options.ExtractText);
            }
            return Task.CompletedTask;
        }
    }
}
