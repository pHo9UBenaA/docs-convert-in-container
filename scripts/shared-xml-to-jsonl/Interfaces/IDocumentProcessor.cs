using System.Threading;
using System.Threading.Tasks;

namespace SharedXmlToJsonl.Interfaces
{
    /// <summary>
    /// Interface for document processors.
    /// </summary>
    public interface IDocumentProcessor
    {
        /// <summary>
        /// Processes a document asynchronously.
        /// </summary>
        /// <param name="inputPath">The path to the input document.</param>
        /// <param name="outputDirectory">The directory to write output files.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The processing result.</returns>
        Task<ProcessingResult> ProcessAsync(string inputPath, string outputDirectory, CancellationToken cancellationToken = default);
    }
}