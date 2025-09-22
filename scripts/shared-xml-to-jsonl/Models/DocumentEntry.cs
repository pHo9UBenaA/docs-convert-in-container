using System;

namespace SharedXmlToJsonl.Models;

/// <summary>
/// Represents a document entry for JSONL output
/// </summary>
public class DocumentEntry
{
    public string Document { get; set; } = string.Empty;
    public string DocumentType { get; set; } = string.Empty;
    public int PageNumber { get; set; }
    public string? SheetName { get; set; }
    public object? Element { get; set; }
}