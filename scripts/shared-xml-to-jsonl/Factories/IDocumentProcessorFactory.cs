namespace SharedXmlToJsonl.Factories
{
    /// <summary>
    /// Factory interface for creating document processors based on file type.
    /// </summary>
    public interface IDocumentProcessorFactory
    {
        /// <summary>
        /// Creates a document processor for the specified file extension.
        /// </summary>
        /// <param name="fileExtension">The file extension (e.g., ".pptx", ".xlsx").</param>
        /// <returns>The appropriate document processor.</returns>
        Interfaces.IDocumentProcessor CreateProcessor(string fileExtension);

        /// <summary>
        /// Determines whether the specified file extension is supported.
        /// </summary>
        /// <param name="fileExtension">The file extension to check.</param>
        /// <returns>True if the file extension is supported; otherwise, false.</returns>
        bool IsSupported(string fileExtension);
    }
}
