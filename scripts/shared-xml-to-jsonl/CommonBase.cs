// File: CommonBase.cs
// Specification: Common base definitions shared between PPTX and XLSX processors

using System.Text;

namespace SharedXmlToJsonl;

/// <summary>
/// Common base class containing shared constants and type definitions
/// </summary>
public static class CommonBase
{
    // Exit codes
    public const int ExitSuccess = 0;
    public const int ExitUsageError = 2;
    public const int ExitProcessingError = 3;

    // UTF-8 encoding without BOM
    public static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
}

/// <summary>
/// Information about a package relationship
/// </summary>
public record RelationshipInfo(string Id, string Type, string Target, string TargetMode);

/// <summary>
/// JSONL entry containing package part information
/// </summary>
public record JsonlEntry(string PartName, string ContentType, IReadOnlyList<RelationshipInfo> Relationships, int SizeBytes, string Xml);

/// <summary>
/// Error information for processing failures
/// </summary>
public record ErrorInfo(string Xml, string Error);