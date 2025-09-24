using System.Threading;
using System.Threading.Tasks;
using SharedXmlToJsonl.Configuration;
using SharedXmlToJsonl.Models;

namespace SharedXmlToJsonl.Interfaces
{
    /// <summary>
    /// Interface for PPTX document processors.
    /// </summary>
    public interface IPptxProcessor : IDocumentProcessor
    {
        Task<ProcessingResult> ProcessAsync(
            string inputPath,
            string outputDirectory,
            ProcessingOptions options,
            CancellationToken cancellationToken = default);
    }
}
