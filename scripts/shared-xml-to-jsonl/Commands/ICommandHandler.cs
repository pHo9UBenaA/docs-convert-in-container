using System;
using System.Threading;
using System.Threading.Tasks;

namespace SharedXmlToJsonl.Commands
{
    /// <summary>
    /// Represents a command handler that can execute operations with specific options.
    /// </summary>
    /// <typeparam name="TOptions">The type of options that this command handler accepts.</typeparam>
    public interface ICommandHandler<TOptions> where TOptions : CommandHandlerOptions
    {
        /// <summary>
        /// Executes the command asynchronously with the provided options.
        /// </summary>
        /// <param name="options">The command options.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The exit code (0 for success, non-zero for failure).</returns>
        Task<int> ExecuteAsync(TOptions options, CancellationToken cancellationToken = default);

        /// <summary>
        /// Sets up the command handler with the service provider.
        /// </summary>
        /// <param name="serviceProvider">The service provider for dependency injection.</param>
        void SetupCommand(IServiceProvider serviceProvider);
    }
}