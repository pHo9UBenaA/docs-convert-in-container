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
using SharedXmlToJsonl;

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

    // Position, Size, Transform are now from SharedXmlToJsonl namespace

    private record SlideElement(
        int SlideNumber,
        string ElementType,
        int ElementIndex,
        string? Text = null,
        string? ShapeId = null,
        string? ShapeName = null,
        Transform? Transform = null,
        string? ShapeType = null,
        int? GroupLevel = null,
        string? ParentGroupId = null,
        CustomGeometry? CustomGeometry = null,
        string? OleObjectType = null,
        string? ContentPartRef = null,
        LineProperties? LineProperties = null,
        bool? HasFill = null,
        string? FillColor = null,
        SlideMetadata? Metadata = null,
        ErrorInfo? ErrorInfo = null);

    [JsonSerializable(typeof(SlideElement))]
    [JsonSerializable(typeof(JsonlEntry))]
    [JsonSerializable(typeof(RelationshipInfo))]
    [JsonSerializable(typeof(IReadOnlyList<RelationshipInfo>))]
    [JsonSerializable(typeof(SlideMetadata))]
    [JsonSerializable(typeof(ErrorInfo))]
    [JsonSerializable(typeof(Position))]
    [JsonSerializable(typeof(Size))]
    [JsonSerializable(typeof(Transform))]
    [JsonSerializable(typeof(CustomGeometry))]
    [JsonSerializable(typeof(TableCell))]
    [JsonSerializable(typeof(LineProperties))]
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

    private static void ProcessGroupShape(XElement groupElement, List<SlideElement> elements, ref int elementIndex,
        int slideNumber, XNamespace p, XNamespace a, int groupLevel, string? parentGroupId)
    {
        // Check if this is actually a group shape (not the root spTree)
        bool isRootSpTree = groupElement.Name == (p + "spTree");

        if (!isRootSpTree)
        {
            // Extract group information
            var grpId = groupElement.Element(p + "nvGrpSpPr")?.Element(p + "cNvPr")?.Attribute("id")?.Value;
            var grpName = groupElement.Element(p + "nvGrpSpPr")?.Element(p + "cNvPr")?.Attribute("name")?.Value;
            var grpSpPr = groupElement.Element(p + "grpSpPr");
            var transform = grpSpPr != null ? ExtractTransform(grpSpPr, a) : null;

            // Add the group itself as an element
            elements.Add(new SlideElement(
                slideNumber,
                "group_shape",
                elementIndex++,
                ShapeId: grpId,
                ShapeName: grpName,
                Transform: transform,
                GroupLevel: groupLevel,
                ParentGroupId: parentGroupId
            ));

            // Update parent ID for children
            parentGroupId = grpId;
        }

        // Process child elements within the group
        foreach (var child in groupElement.Elements())
        {
            var childName = child.Name.LocalName;

            if (childName == "sp") // Regular shape
            {
                ProcessShape(child, elements, ref elementIndex, slideNumber, p, a, groupLevel + 1, parentGroupId);
            }
            else if (childName == "grpSp") // Nested group
            {
                ProcessGroupShape(child, elements, ref elementIndex, slideNumber, p, a, groupLevel + 1, parentGroupId);
            }
            else if (childName == "pic") // Picture
            {
                ProcessPicture(child, elements, ref elementIndex, slideNumber, p, a, groupLevel + 1, parentGroupId);
            }
            else if (childName == "cxnSp") // Connector
            {
                ProcessConnector(child, elements, ref elementIndex, slideNumber, p, a, groupLevel + 1, parentGroupId);
            }
            else if (childName == "graphicFrame") // Graphic frame (table, chart, etc.)
            {
                ProcessGraphicFrame(child, elements, ref elementIndex, slideNumber, p, a, groupLevel + 1, parentGroupId);
            }
            else if (childName == "contentPart") // Content part
            {
                ProcessContentPart(child, elements, ref elementIndex, slideNumber, p, a, groupLevel + 1, parentGroupId);
            }
        }
    }

    private static void ProcessShape(XElement shape, List<SlideElement> elements, ref int elementIndex,
        int slideNumber, XNamespace p, XNamespace a, int groupLevel, string? parentGroupId)
    {
        var shapeId = shape.Element(p + "nvSpPr")?.Element(p + "cNvPr")?.Attribute("id")?.Value;
        var shapeName = shape.Element(p + "nvSpPr")?.Element(p + "cNvPr")?.Attribute("name")?.Value;

        var spPr = shape.Element(p + "spPr");
        var transform = spPr != null ? ExtractTransform(spPr, a) : null;

        // Check for custom geometry
        var custGeom = spPr?.Element(a + "custGeom");
        var customGeometry = XmlUtilities.ExtractCustomGeometry(custGeom, a);

        // Extract shape type (preset geometry)
        var prstGeom = spPr?.Element(a + "prstGeom");
        var shapeType = prstGeom?.Attribute("prst")?.Value ?? (customGeometry != null ? "custom" : null);

        // Extract line properties
        var lineProperties = ExtractLineProperties(spPr, a);

        // Extract fill information
        var (hasFill, fillColor) = ExtractFillInfo(spPr, a);

        // First, record the shape itself
        elements.Add(new SlideElement(
            slideNumber,
            "shape",
            elementIndex++,
            ShapeId: shapeId,
            ShapeName: shapeName,
            Transform: transform,
            ShapeType: shapeType,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId,
            CustomGeometry: customGeometry,
            LineProperties: lineProperties,
            HasFill: hasFill,
            FillColor: fillColor
        ));

        // Then, extract and record text as separate elements
        var paragraphs = shape.Descendants(a + "p");
        foreach (var paragraph in paragraphs)
        {
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
                    ShapeName: shapeName,
                    Transform: transform,
                    GroupLevel: groupLevel,
                    ParentGroupId: parentGroupId
                ));
            }
        }

        // Check if the shape is purely a line (no fill, has line)
        if (lineProperties != null && hasFill == false && shapeType == "line")
        {
            elements.Add(new SlideElement(
                slideNumber,
                "line",
                elementIndex++,
                ShapeId: shapeId,
                ShapeName: shapeName,
                Transform: transform,
                GroupLevel: groupLevel,
                ParentGroupId: parentGroupId,
                LineProperties: lineProperties
            ));
        }
    }

    private static void ProcessPicture(XElement pic, List<SlideElement> elements, ref int elementIndex,
        int slideNumber, XNamespace p, XNamespace a, int groupLevel, string? parentGroupId)
    {
        var picName = pic.Element(p + "nvPicPr")?.Element(p + "cNvPr")?.Attribute("name")?.Value;
        var picId = pic.Element(p + "nvPicPr")?.Element(p + "cNvPr")?.Attribute("id")?.Value;

        var spPr = pic.Element(p + "spPr");
        var transform = spPr != null ? ExtractTransform(spPr, a) : null;

        // Extract line properties for picture border
        var lineProperties = ExtractLineProperties(spPr, a);

        elements.Add(new SlideElement(
            slideNumber,
            "image",
            elementIndex++,
            ShapeId: picId,
            ShapeName: picName,
            Transform: transform,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId,
            LineProperties: lineProperties
        ));
    }

    private static void ProcessConnector(XElement cxnSp, List<SlideElement> elements, ref int elementIndex,
        int slideNumber, XNamespace p, XNamespace a, int groupLevel, string? parentGroupId)
    {
        var connId = cxnSp.Element(p + "nvCxnSpPr")?.Element(p + "cNvPr")?.Attribute("id")?.Value;
        var connName = cxnSp.Element(p + "nvCxnSpPr")?.Element(p + "cNvPr")?.Attribute("name")?.Value;

        var spPr = cxnSp.Element(p + "spPr");
        var transform = spPr != null ? ExtractTransform(spPr, a) : null;

        var prstGeom = spPr?.Element(a + "prstGeom");
        var shapeType = prstGeom?.Attribute("prst")?.Value;

        // Extract line properties for connector
        var lineProperties = ExtractLineProperties(spPr, a);

        elements.Add(new SlideElement(
            slideNumber,
            "connector",
            elementIndex++,
            ShapeId: connId,
            ShapeName: connName,
            Transform: transform,
            ShapeType: shapeType,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId,
            LineProperties: lineProperties
        ));
    }

    private static void ProcessGraphicFrame(XElement graphicFrame, List<SlideElement> elements, ref int elementIndex,
        int slideNumber, XNamespace p, XNamespace a, int groupLevel, string? parentGroupId)
    {
        var frameId = graphicFrame.Element(p + "nvGraphicFramePr")?.Element(p + "cNvPr")?.Attribute("id")?.Value;
        var frameName = graphicFrame.Element(p + "nvGraphicFramePr")?.Element(p + "cNvPr")?.Attribute("name")?.Value;

        var xfrm = graphicFrame.Element(p + "xfrm");
        var transform = xfrm != null ? ExtractTransform(xfrm, a) : null;

        // Check if it contains a table
        var graphicData = graphicFrame.Descendants(a + "graphicData").FirstOrDefault();
        if (graphicData != null)
        {
            var table = graphicData.Element(a + "tbl");
            if (table != null)
            {
                ProcessTable(table, elements, ref elementIndex, slideNumber, a, frameId, frameName, transform, groupLevel, parentGroupId);
                return;
            }

            // Check for SmartArt
            XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            var relIds = graphicData.Element(dgm + "relIds");
            if (relIds != null)
            {
                elements.Add(new SlideElement(
                    slideNumber,
                    "smartart",
                    elementIndex++,
                    ShapeId: frameId,
                    ShapeName: frameName,
                    Transform: transform,
                    GroupLevel: groupLevel,
                    ParentGroupId: parentGroupId
                ));
                return;
            }

            // Check for OLE objects
            XNamespace o = "urn:schemas-microsoft-com:office:office";
            var oleObj = graphicData.Descendants(o + "oleObj").FirstOrDefault();
            if (oleObj != null)
            {
                var progId = oleObj.Attribute("progId")?.Value;
                elements.Add(new SlideElement(
                    slideNumber,
                    "ole_object",
                    elementIndex++,
                    ShapeId: frameId,
                    ShapeName: frameName,
                    Transform: transform,
                    GroupLevel: groupLevel,
                    ParentGroupId: parentGroupId,
                    OleObjectType: progId
                ));
                return;
            }
        }

        // Generic graphic frame if type not identified
        elements.Add(new SlideElement(
            slideNumber,
            "graphic_frame",
            elementIndex++,
            ShapeId: frameId,
            ShapeName: frameName,
            Transform: transform,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId
        ));
    }

    private static void ProcessTable(XElement table, List<SlideElement> elements, ref int elementIndex,
        int slideNumber, XNamespace a, string? tableId, string? tableName, Transform? transform, int groupLevel, string? parentGroupId)
    {
        // Add table element
        elements.Add(new SlideElement(
            slideNumber,
            "table",
            elementIndex++,
            ShapeId: tableId,
            ShapeName: tableName,
            Transform: transform,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId
        ));

        // Process table rows
        var rows = table.Elements(a + "tr");
        int rowIndex = 0;
        foreach (var row in rows)
        {
            var cells = row.Elements(a + "tc");
            int colIndex = 0;
            foreach (var cell in cells)
            {
                // Extract text from cell
                var cellTexts = cell.Descendants(a + "t").Select(t => t.Value);
                var cellText = string.Join(" ", cellTexts);

                if (!string.IsNullOrWhiteSpace(cellText))
                {
                    elements.Add(new SlideElement(
                        slideNumber,
                        "table_cell",
                        elementIndex++,
                        Text: cellText,
                        ShapeId: $"{tableId}_R{rowIndex}C{colIndex}",
                        ShapeName: $"Cell[{rowIndex},{colIndex}]",
                        GroupLevel: groupLevel + 1,
                        ParentGroupId: tableId
                    ));
                }
                colIndex++;
            }
            rowIndex++;
        }
    }

    private static void ProcessContentPart(XElement contentPart, List<SlideElement> elements, ref int elementIndex,
        int slideNumber, XNamespace p, XNamespace a, int groupLevel, string? parentGroupId)
    {
        // Extract content part reference
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var relId = contentPart.Attribute(r + "id")?.Value;

        elements.Add(new SlideElement(
            slideNumber,
            "content_part",
            elementIndex++,
            ContentPartRef: relId,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId
        ));
    }

    private static Transform? ExtractTransform(XElement element, XNamespace a)
    {
        var xfrm = element.Element(a + "xfrm");
        return XmlUtilities.ExtractTransformFromXfrm(xfrm, a);
    }

    private static LineProperties? ExtractLineProperties(XElement? spPr, XNamespace a)
    {
        if (spPr == null) return null;

        var ln = spPr.Element(a + "ln");
        if (ln == null) return null;

        // Extract line width
        var width = ln.Attribute("w")?.Value;
        long? lineWidth = null;
        if (width != null && long.TryParse(width, out var w))
        {
            lineWidth = w;
        }

        // Extract line color
        string? lineColor = null;
        var solidFill = ln.Element(a + "solidFill");
        if (solidFill != null)
        {
            var srgbClr = solidFill.Element(a + "srgbClr");
            if (srgbClr != null)
            {
                lineColor = srgbClr.Attribute("val")?.Value;
            }
        }

        // Extract dash style
        var prstDash = ln.Element(a + "prstDash");
        var dashStyle = prstDash?.Attribute("val")?.Value;

        // Extract compound line type
        var cmpd = ln.Attribute("cmpd")?.Value;

        if (lineWidth == null && lineColor == null && dashStyle == null && cmpd == null)
        {
            return null;
        }

        return new LineProperties(
            Color: lineColor,
            Width: lineWidth,
            DashStyle: dashStyle,
            CompoundLineType: cmpd
        );
    }

    private static (bool? hasFill, string? fillColor) ExtractFillInfo(XElement? spPr, XNamespace a)
    {
        if (spPr == null) return (null, null);

        // Check for no fill
        var noFill = spPr.Element(a + "noFill");
        if (noFill != null)
        {
            return (false, null);
        }

        // Check for solid fill
        var solidFill = spPr.Element(a + "solidFill");
        if (solidFill != null)
        {
            var srgbClr = solidFill.Element(a + "srgbClr");
            var fillColor = srgbClr?.Attribute("val")?.Value;
            return (true, fillColor);
        }

        // Check for gradient fill
        var gradFill = spPr.Element(a + "gradFill");
        if (gradFill != null)
        {
            return (true, "gradient");
        }

        // Check for pattern fill
        var pattFill = spPr.Element(a + "pattFill");
        if (pattFill != null)
        {
            return (true, "pattern");
        }

        return (null, null);
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

            // Process the root shape tree as a group (this will recursively process all elements)
            var rootSpTree = doc.Descendants(p + "spTree").FirstOrDefault();
            if (rootSpTree != null)
            {
                ProcessGroupShape(rootSpTree, elements, ref elementIndex, slideNumber, p, a, 0, null);
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
