using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SharedXmlToJsonl.Interfaces
{
    /// <summary>
    /// Interface for writing JSON data.
    /// </summary>
    public interface IJsonWriter
    {
        /// <summary>
        /// Writes a single JSON line asynchronously.
        /// </summary>
        /// <typeparam name="T">The type of object to serialize.</typeparam>
        /// <param name="writer">The stream writer.</param>
        /// <param name="obj">The object to serialize.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        Task WriteJsonLineAsync<T>(StreamWriter writer, T obj, CancellationToken cancellationToken = default);

        /// <summary>
        /// Writes multiple JSON lines to a file asynchronously.
        /// </summary>
        /// <typeparam name="T">The type of objects to serialize.</typeparam>
        /// <param name="filePath">The path to the output file.</param>
        /// <param name="objects">The objects to serialize.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        Task WriteJsonLinesAsync<T>(string filePath, IEnumerable<T> objects, CancellationToken cancellationToken = default);
    }
}