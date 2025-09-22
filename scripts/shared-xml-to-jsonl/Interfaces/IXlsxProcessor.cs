using System.Threading;
using System.Threading.Tasks;
using SharedXmlToJsonl.Configuration;

namespace SharedXmlToJsonl.Interfaces
{
    /// <summary>
    /// Interface for XLSX document processors.
    /// </summary>
    public interface IXlsxProcessor : IDocumentProcessor
    {
        Task<ProcessingResult> ProcessAsync(
            string inputPath,
            string outputDirectory,
            ProcessingOptions options,
            CancellationToken cancellationToken = default);
    }
}