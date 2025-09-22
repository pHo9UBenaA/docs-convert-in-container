using System;

namespace SharedXmlToJsonl.Exceptions;

public class DocumentProcessingException : Exception
{
    public string DocumentPath { get; }
    public ProcessingStage Stage { get; }
    public int? PageNumber { get; }
    public string? ElementId { get; }

    public DocumentProcessingException(
        string message,
        string documentPath,
        ProcessingStage stage,
        Exception? innerException = null)
        : base(message, innerException)
    {
        DocumentPath = documentPath ?? throw new ArgumentNullException(nameof(documentPath));
        Stage = stage;
    }

    public DocumentProcessingException(
        string message,
        string documentPath,
        ProcessingStage stage,
        int pageNumber,
        Exception? innerException = null)
        : this(message, documentPath, stage, innerException)
    {
        PageNumber = pageNumber;
    }

    public DocumentProcessingException(
        string message,
        string documentPath,
        ProcessingStage stage,
        string elementId,
        Exception? innerException = null)
        : this(message, documentPath, stage, innerException)
    {
        ElementId = elementId;
    }
}

public enum ProcessingStage
{
    Opening,
    Parsing,
    Transforming,
    Writing,
    Validating,
    Extracting
}