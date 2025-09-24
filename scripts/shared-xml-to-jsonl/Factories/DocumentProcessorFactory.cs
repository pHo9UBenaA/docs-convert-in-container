using System;
using System.Collections.Generic;
using Microsoft.Extensions.DependencyInjection;
using SharedXmlToJsonl.Interfaces;

namespace SharedXmlToJsonl.Factories
{
    /// <summary>
    /// Factory implementation for creating document processors.
    /// </summary>
    public class DocumentProcessorFactory : IDocumentProcessorFactory
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly Dictionary<string, Type> _processorTypes;

        /// <summary>
        /// Initializes a new instance of the DocumentProcessorFactory class.
        /// </summary>
        /// <param name="serviceProvider">The service provider for dependency injection.</param>
        public DocumentProcessorFactory(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
            _processorTypes = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase)
            {
                [".pptx"] = typeof(IPptxProcessor),
                [".xlsx"] = typeof(IXlsxProcessor)
            };
        }

        /// <summary>
        /// Creates a document processor for the specified file extension.
        /// </summary>
        public IDocumentProcessor CreateProcessor(string fileExtension)
        {
            if (!_processorTypes.TryGetValue(fileExtension, out var processorType))
            {
                throw new NotSupportedException($"File extension {fileExtension} is not supported. Supported extensions: {string.Join(", ", _processorTypes.Keys)}");
            }

            var processor = _serviceProvider.GetService(processorType) as IDocumentProcessor;
            if (processor == null)
            {
                throw new InvalidOperationException($"No processor registered for type {processorType.Name}");
            }

            return processor;
        }

        /// <summary>
        /// Determines whether the specified file extension is supported.
        /// </summary>
        public bool IsSupported(string fileExtension)
        {
            return _processorTypes.ContainsKey(fileExtension);
        }
    }
}
