using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Commands
{
    /// <summary>
    /// Base class for command handlers implementing the Template Method pattern.
    /// </summary>
    /// <typeparam name="TOptions">The type of options for this command handler.</typeparam>
    public abstract partial class CommandHandlerBase<TOptions> : ICommandHandler<TOptions>
        where TOptions : CommandHandlerOptions
    {
        protected ILogger<CommandHandlerBase<TOptions>> Logger { get; }
        protected IServiceProvider ServiceProvider { get; }

        /// <summary>
        /// Initializes a new instance of the CommandHandlerBase class.
        /// </summary>
        /// <param name="logger">The logger instance.</param>
        /// <param name="serviceProvider">The service provider for dependency injection.</param>
        protected CommandHandlerBase(
            ILogger<CommandHandlerBase<TOptions>> logger,
            IServiceProvider serviceProvider)
        {
            ArgumentNullException.ThrowIfNull(logger);
            Logger = logger;
            ArgumentNullException.ThrowIfNull(serviceProvider);
            ServiceProvider = serviceProvider;
        }

        /// <summary>
        /// Executes the command with error handling and lifecycle management.
        /// </summary>
        public async Task<int> ExecuteAsync(TOptions options, CancellationToken cancellationToken = default)
        {
            try
            {
                // 1. Option validation
                var validationResult = options.Validate();
                if (!validationResult.IsValid)
                {
                    foreach (var error in validationResult.Errors)
                    {
                        LogValidationFailed(Logger, error);
                    }
                    return CommonBase.ExitUsageError;
                }

                // 2. Pre-processing
                await OnBeforeExecuteAsync(options, cancellationToken).ConfigureAwait(false);

                // 3. Main processing (implemented in derived classes)
                var result = await ExecuteCoreAsync(options, cancellationToken).ConfigureAwait(false);

                // 4. Post-processing
                await OnAfterExecuteAsync(options, result, cancellationToken).ConfigureAwait(false);

                return result;
            }
            catch (OperationCanceledException)
            {
                LogOperationCancelled(Logger);
                return CommonBase.ExitProcessingError;
            }
            catch (Exception ex)
            {
                LogCommandExecutionFailed(Logger, ex);
                return CommonBase.ExitProcessingError;
            }
        }

        /// <summary>
        /// Core execution logic to be implemented by derived classes.
        /// </summary>
        /// <param name="options">The command options.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The exit code.</returns>
        protected abstract Task<int> ExecuteCoreAsync(TOptions options, CancellationToken cancellationToken);

        /// <summary>
        /// Hook method called before execution. Can be overridden by derived classes.
        /// </summary>
        protected virtual Task OnBeforeExecuteAsync(TOptions options, CancellationToken cancellationToken)
            => Task.CompletedTask;

        /// <summary>
        /// Hook method called after execution. Can be overridden by derived classes.
        /// </summary>
        protected virtual Task OnAfterExecuteAsync(TOptions options, int result, CancellationToken cancellationToken)
            => Task.CompletedTask;

        /// <summary>
        /// Sets up the command handler with the service provider.
        /// </summary>
        public abstract void SetupCommand(IServiceProvider serviceProvider);

        [LoggerMessage(
            EventId = 3001,
            Level = LogLevel.Error,
            Message = "Validation failed: {error}")]
        private static partial void LogValidationFailed(
            ILogger logger, string error);

        [LoggerMessage(
            EventId = 3002,
            Level = LogLevel.Warning,
            Message = "Operation was cancelled")]
        private static partial void LogOperationCancelled(
            ILogger logger);

        [LoggerMessage(
            EventId = 3003,
            Level = LogLevel.Error,
            Message = "Command execution failed")]
        private static partial void LogCommandExecutionFailed(
            ILogger logger, Exception ex);
    }
}
