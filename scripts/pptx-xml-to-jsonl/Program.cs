// File: Program.cs
// Specification: CLI to transform PPTX XML package parts into JSONL preserving raw XML and relationships.

using System.Globalization;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace PptxXmlToJsonl;

internal static partial class Program
{
    private const int ExitSuccess = 0;
    private const int ExitUsageError = 2;
    private const int ExitProcessingError = 3;

    private const string RelationshipContentType = "application/vnd.openxmlformats-package.relationships+xml";
    private const string DefaultXmlContentType = "application/xml";

    // Use custom context with relaxed encoding
    private static readonly SourceGenerationContext JsonContext = SourceGenerationContext.Custom;

    private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    private record RelationshipInfo(string Id, string Type, string Target, string TargetMode);

    private record JsonlEntry(string PartName, string ContentType, IReadOnlyList<RelationshipInfo> Relationships, int SizeBytes, string Xml);

    private record SlideMetadata(int SlideNumber);
    private record ErrorInfo(string Xml, string Error);

    private record SlideElement(
        int SlideNumber,
        string ElementType,
        int ElementIndex,
        string? Text = null,
        string? ShapeId = null,
        string? ShapeName = null,
        SlideMetadata? Metadata = null,
        ErrorInfo? ErrorInfo = null);

    [JsonSerializable(typeof(SlideElement))]
    [JsonSerializable(typeof(JsonlEntry))]
    [JsonSerializable(typeof(RelationshipInfo))]
    [JsonSerializable(typeof(IReadOnlyList<RelationshipInfo>))]
    [JsonSerializable(typeof(SlideMetadata))]
    [JsonSerializable(typeof(ErrorInfo))]
    private partial class SourceGenerationContext : JsonSerializerContext
    {
        private static readonly JsonSerializerOptions _options = new()
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            WriteIndented = false,
            PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower
        };

        public static SourceGenerationContext Custom { get; } = new SourceGenerationContext(_options);
    }

    private static int Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.Error.WriteLine("Usage: pptx-xml-to-jsonl <input.pptx> <output-directory>");
            return ExitUsageError;
        }

        var inputPath = Path.GetFullPath(args[0]);
        var outputDirectory = Path.GetFullPath(args[1]);

        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return ExitProcessingError;
        }

        try
        {
            var entries = ExtractEntries(inputPath);
            WritePerSlideJsonLines(inputPath, outputDirectory, entries);
            return ExitSuccess;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return ExitProcessingError;
        }
    }

    private static IReadOnlyList<JsonlEntry> ExtractEntries(string pptxPath)
    {
        using var package = Package.Open(pptxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = ZipFile.OpenRead(pptxPath);

        var entries = archive.Entries
            .Where(IsXmlEntry)
            .Select(entry => BuildEntry(entry, package))
            .OrderBy(entry => entry.PartName, StringComparer.Ordinal)
            .ToList();

        return entries;
    }

    private static void WritePerSlideJsonLines(string pptxPath, string outputDirectory, IReadOnlyList<JsonlEntry> entries)
    {
        Directory.CreateDirectory(outputDirectory);

        var baseName = Path.GetFileNameWithoutExtension(pptxPath);
        var slideNumbers = GetSlideNumbers(entries);

        foreach (var slideNumber in slideNumbers)
        {
            var filePath = Path.Combine(outputDirectory, $"{baseName}_page-{slideNumber}.jsonl");
            var slideSpecificEntries = FilterEntriesForSlide(entries, slideNumber);
            WriteSlideElementsAsJsonLines(filePath, slideSpecificEntries, slideNumber);
        }
    }

    private static IReadOnlyList<int> GetSlideNumbers(IReadOnlyList<JsonlEntry> entries)
    {
        var slideNumbers = new SortedSet<int>();

        foreach (var entry in entries)
        {
            var slideNumber = TryExtractSlideNumber(entry.PartName);
            if (slideNumber.HasValue)
            {
                slideNumbers.Add(slideNumber.Value);
            }
        }

        if (slideNumbers.Count == 0)
        {
            return new[] { 1 };
        }

        return slideNumbers.ToArray();
    }

    private static IReadOnlyList<JsonlEntry> FilterEntriesForSlide(IReadOnlyList<JsonlEntry> entries, int slideNumber)
    {
        var filteredEntries = new List<JsonlEntry>();

        foreach (var entry in entries)
        {
            if (IsSlideRelatedEntry(entry.PartName, slideNumber))
            {
                filteredEntries.Add(entry);
            }
        }

        return filteredEntries;
    }

    private static bool IsSlideRelatedEntry(string partName, int slideNumber)
    {
        // Check if the entry is directly related to the specified slide
        var slidePartName = $"/ppt/slides/slide{slideNumber}.xml";
        var slideRelsPartName = $"/ppt/slides/_rels/slide{slideNumber}.xml.rels";

        // Only include slide-specific entries
        return partName.Equals(slidePartName, StringComparison.OrdinalIgnoreCase) ||
               partName.Equals(slideRelsPartName, StringComparison.OrdinalIgnoreCase);
    }

    private static JsonlEntry BuildEntry(ZipArchiveEntry entry, Package package)
    {
        var partName = "/" + entry.FullName.Replace('\\', '/');

        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Utf8NoBom, detectEncodingFromByteOrderMarks: true);
        var xml = reader.ReadToEnd();
        var sizeBytes = Utf8NoBom.GetByteCount(xml);

        if (TryGetPackagePart(package, partName, out var packagePart))
        {
            if (string.Equals(packagePart.ContentType, RelationshipContentType, StringComparison.OrdinalIgnoreCase))
            {
                return new JsonlEntry(partName, packagePart.ContentType, Array.Empty<RelationshipInfo>(), sizeBytes, xml);
            }

            var relationships = packagePart
                .GetRelationships()
                .Select(rel => new RelationshipInfo(
                    rel.Id,
                    rel.RelationshipType,
                    rel.TargetUri?.ToString() ?? string.Empty,
                    rel.TargetMode.ToString()))
                .OrderBy(rel => rel.Id, StringComparer.Ordinal)
                .ToArray();

            return new JsonlEntry(partName, packagePart.ContentType, relationships, sizeBytes, xml);
        }

        var fallbackContentType = DetermineFallbackContentType(partName);
        return new JsonlEntry(partName, fallbackContentType, Array.Empty<RelationshipInfo>(), sizeBytes, xml);
    }

    private static bool TryGetPackagePart(Package package, string partName, out PackagePart packagePart)
    {
        packagePart = null!;
        try
        {
            var uri = PackUriHelper.CreatePartUri(new Uri(partName, UriKind.RelativeOrAbsolute));
            if (package.PartExists(uri))
            {
                packagePart = package.GetPart(uri);
                return true;
            }
        }
        catch (ArgumentException)
        {
            // Relationship parts such as "/_rels/.rels" or invalid URIs fall through here.
        }
        return false;
    }

    private static string DetermineFallbackContentType(string partName)
    {
        if (partName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
        {
            return RelationshipContentType;
        }

        return DefaultXmlContentType;
    }

    private static bool IsXmlEntry(ZipArchiveEntry entry)
    {
        var name = entry.FullName;
        if (name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
    }

    private static int? TryExtractSlideNumber(string partName)
    {
        const string slidePrefix = "/ppt/slides/slide";
        if (!partName.StartsWith(slidePrefix, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var suffix = partName.Substring(slidePrefix.Length);
        var digits = new string(suffix.TakeWhile(char.IsDigit).ToArray());
        if (digits.Length == 0)
        {
            return null;
        }

        if (int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out var value))
        {
            return value;
        }

        return null;
    }

    private static void WriteSlideElementsAsJsonLines(string outputPath, IReadOnlyList<JsonlEntry> entries, int slideNumber)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var stream = File.Open(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, Utf8NoBom);

        foreach (var entry in entries)
        {
            // Check if this is a slide XML file
            if (entry.PartName.StartsWith("/ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                entry.PartName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !entry.PartName.Contains(".rels", StringComparison.OrdinalIgnoreCase))
            {
                // Parse slide XML and write elements separately
                var elements = ExtractSlideElements(entry.Xml, slideNumber);
                foreach (var element in elements)
                {
                    writer.WriteLine(JsonSerializer.Serialize(element, JsonContext.SlideElement));
                }
            }
            // Skip relationship files - they don't contain useful content for understanding the page
        }
    }

    private static IReadOnlyList<SlideElement> ExtractSlideElements(string xml, int slideNumber)
    {
        var elements = new List<SlideElement>();
        var elementIndex = 0;

        try
        {
            var doc = XDocument.Parse(xml);
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            // Add slide metadata as first element
            elements.Add(new SlideElement(
                slideNumber,
                "slide_metadata",
                elementIndex++,
                Metadata: new SlideMetadata(slideNumber)
            ));

            // Extract all shapes with text
            var shapes = doc.Descendants(p + "sp");
            foreach (var shape in shapes)
            {
                var shapeId = shape.Element(p + "nvSpPr")?.Element(p + "cNvPr")?.Attribute("id")?.Value;
                var shapeName = shape.Element(p + "nvSpPr")?.Element(p + "cNvPr")?.Attribute("name")?.Value;

                // Extract paragraphs from the shape
                var paragraphs = shape.Descendants(a + "p");
                foreach (var paragraph in paragraphs)
                {
                    // Extract text runs from the paragraph
                    var textRuns = paragraph.Descendants(a + "t");
                    var paragraphText = string.Join("", textRuns.Select(t => t.Value));

                    if (!string.IsNullOrWhiteSpace(paragraphText))
                    {
                        elements.Add(new SlideElement(
                            slideNumber,
                            "text",
                            elementIndex++,
                            Text: paragraphText,
                            ShapeId: shapeId,
                            ShapeName: shapeName
                        ));
                    }
                }

                // If shape has no text, still record it
                if (!shape.Descendants(a + "t").Any())
                {
                    elements.Add(new SlideElement(
                        slideNumber,
                        "shape",
                        elementIndex++,
                        ShapeId: shapeId,
                        ShapeName: shapeName
                    ));
                }
            }

            // Extract images/media references
            var pics = doc.Descendants(p + "pic");
            foreach (var pic in pics)
            {
                var picName = pic.Element(p + "nvPicPr")?.Element(p + "cNvPr")?.Attribute("name")?.Value;
                var picId = pic.Element(p + "nvPicPr")?.Element(p + "cNvPr")?.Attribute("id")?.Value;

                elements.Add(new SlideElement(
                    slideNumber,
                    "image",
                    elementIndex++,
                    ShapeId: picId,
                    ShapeName: picName
                ));
            }
        }
        catch (Exception ex)
        {
            // If XML parsing fails, return slide as single element with raw XML
            elements.Add(new SlideElement(
                slideNumber,
                "raw_xml",
                0,
                ErrorInfo: new ErrorInfo(xml, ex.Message)
            ));
        }

        return elements;
    }

    private static void WriteJsonLines(string outputPath, IReadOnlyList<JsonlEntry> entries)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var stream = File.Open(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, Utf8NoBom);

        foreach (var entry in entries)
        {
            writer.WriteLine(JsonSerializer.Serialize(entry, JsonContext.JsonlEntry));
        }
    }
}
