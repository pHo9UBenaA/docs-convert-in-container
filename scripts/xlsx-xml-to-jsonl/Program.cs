// File: Program.cs
// Specification: CLI to transform XLSX XML package parts into JSONL preserving cell data and relationships.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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

namespace XlsxXmlToJsonl;

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

    private record SheetMetadata(int SheetNumber, string SheetName);
    private record ErrorInfo(string Xml, string Error);
    // Position, Size, Transform, CellAnchor are now from SharedXmlToJsonl namespace

    private record SheetElement(
        int SheetNumber,
        string SheetName,
        string ElementType,
        int ElementIndex,
        string? CellReference = null,
        string? Value = null,
        string? Formula = null,
        string? DataType = null,
        int? Row = null,
        int? Column = null,
        Transform? Transform = null,
        string? ShapeType = null,
        string? ShapeId = null,
        string? ShapeName = null,
        int? GroupLevel = null,
        string? ParentGroupId = null,
        CustomGeometry? CustomGeometry = null,
        string? OleObjectType = null,
        SheetMetadata? Metadata = null,
        ErrorInfo? ErrorInfo = null);

    [JsonSerializable(typeof(SheetElement))]
    [JsonSerializable(typeof(JsonlEntry))]
    [JsonSerializable(typeof(RelationshipInfo))]
    [JsonSerializable(typeof(IReadOnlyList<RelationshipInfo>))]
    [JsonSerializable(typeof(SheetMetadata))]
    [JsonSerializable(typeof(ErrorInfo))]
    [JsonSerializable(typeof(Position))]
    [JsonSerializable(typeof(Size))]
    [JsonSerializable(typeof(CellAnchor))]
    [JsonSerializable(typeof(Transform))]
    [JsonSerializable(typeof(CustomGeometry))]
    [JsonSerializable(typeof(TableCell))]
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
            Console.Error.WriteLine("Usage: xlsx-xml-to-jsonl <input.xlsx> <output-directory>");
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
            WritePerSheetJsonLines(inputPath, outputDirectory, entries);
            return ExitSuccess;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return ExitProcessingError;
        }
    }

    private static IReadOnlyList<JsonlEntry> ExtractEntries(string xlsxPath)
    {
        using var package = Package.Open(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = ZipFile.OpenRead(xlsxPath);

        var entries = archive.Entries
            .Where(IsXmlEntry)
            .Select(entry => BuildEntry(entry, package))
            .OrderBy(entry => entry.PartName, StringComparer.Ordinal)
            .ToList();

        return entries;
    }

    private static void WritePerSheetJsonLines(string xlsxPath, string outputDirectory, IReadOnlyList<JsonlEntry> entries)
    {
        Directory.CreateDirectory(outputDirectory);

        var baseName = Path.GetFileNameWithoutExtension(xlsxPath);
        var sheetInfos = GetSheetInfos(entries);

        // Get shared strings for string lookups
        var sharedStrings = ExtractSharedStrings(entries);

        foreach (var (sheetNumber, sheetName) in sheetInfos)
        {
            var filePath = Path.Combine(outputDirectory, $"{baseName}_sheet{sheetNumber}.jsonl");
            var sheetSpecificEntries = FilterEntriesForSheet(entries, sheetNumber);
            WriteSheetElementsAsJsonLines(filePath, sheetSpecificEntries, sheetNumber, sheetName, sharedStrings);
        }
    }

    private static IReadOnlyList<(int SheetNumber, string SheetName)> GetSheetInfos(IReadOnlyList<JsonlEntry> entries)
    {
        var sheetInfos = new List<(int, string)>();

        // Find workbook entry to get sheet names
        var workbookEntry = entries.FirstOrDefault(e => e.PartName == "/xl/workbook.xml");
        if (workbookEntry != null)
        {
            try
            {
                var doc = XDocument.Parse(workbookEntry.Xml);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                var sheets = doc.Descendants(ns + "sheet");
                int sheetNumber = 1;
                foreach (var sheet in sheets)
                {
                    var name = sheet.Attribute("name")?.Value ?? $"Sheet{sheetNumber}";
                    sheetInfos.Add((sheetNumber, name));
                    sheetNumber++;
                }
            }
            catch
            {
                // If we can't parse workbook, fall back to finding sheet files
            }
        }

        // If no sheets found in workbook, look for sheet files directly
        if (sheetInfos.Count == 0)
        {
            var sheetNumbers = new SortedSet<int>();
            foreach (var entry in entries)
            {
                var sheetNumber = TryExtractSheetNumber(entry.PartName);
                if (sheetNumber.HasValue)
                {
                    sheetNumbers.Add(sheetNumber.Value);
                }
            }

            foreach (var num in sheetNumbers)
            {
                sheetInfos.Add((num, $"Sheet{num}"));
            }
        }

        if (sheetInfos.Count == 0)
        {
            return new[] { (1, "Sheet1") };
        }

        return sheetInfos;
    }

    private static IReadOnlyList<string> ExtractSharedStrings(IReadOnlyList<JsonlEntry> entries)
    {
        var sharedStrings = new List<string>();
        var sharedStringsEntry = entries.FirstOrDefault(e => e.PartName == "/xl/sharedStrings.xml");

        if (sharedStringsEntry != null)
        {
            try
            {
                var doc = XDocument.Parse(sharedStringsEntry.Xml);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                var stringItems = doc.Descendants(ns + "si");
                foreach (var si in stringItems)
                {
                    var textElements = si.Descendants(ns + "t");
                    var text = string.Join("", textElements.Select(t => t.Value));
                    sharedStrings.Add(text);
                }
            }
            catch
            {
                // If parsing fails, return empty list
            }
        }

        return sharedStrings;
    }

    private static IReadOnlyList<JsonlEntry> FilterEntriesForSheet(IReadOnlyList<JsonlEntry> entries, int sheetNumber)
    {
        var filteredEntries = new List<JsonlEntry>();

        foreach (var entry in entries)
        {
            if (IsSheetRelatedEntry(entry.PartName, sheetNumber))
            {
                filteredEntries.Add(entry);
            }
        }

        return filteredEntries;
    }

    private static bool IsSheetRelatedEntry(string partName, int sheetNumber)
    {
        // Check if the entry is directly related to the specified sheet
        var sheetPartName = $"/xl/worksheets/sheet{sheetNumber}.xml";
        var sheetRelsPartName = $"/xl/worksheets/_rels/sheet{sheetNumber}.xml.rels";
        var drawingPartName = $"/xl/drawings/drawing{sheetNumber}.xml";

        // Include sheet-specific entries and associated drawings
        return partName.Equals(sheetPartName, StringComparison.OrdinalIgnoreCase) ||
               partName.Equals(sheetRelsPartName, StringComparison.OrdinalIgnoreCase) ||
               partName.Equals(drawingPartName, StringComparison.OrdinalIgnoreCase);
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

    private static int? TryExtractSheetNumber(string partName)
    {
        const string sheetPrefix = "/xl/worksheets/sheet";
        if (!partName.StartsWith(sheetPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var suffix = partName.Substring(sheetPrefix.Length);
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

    private static void WriteSheetElementsAsJsonLines(string outputPath, IReadOnlyList<JsonlEntry> entries, int sheetNumber, string sheetName, IReadOnlyList<string> sharedStrings)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var stream = File.Open(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, Utf8NoBom);

        var allElements = new List<SheetElement>();

        foreach (var entry in entries)
        {
            // Check if this is a sheet XML file
            if (entry.PartName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
                entry.PartName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !entry.PartName.Contains(".rels", StringComparison.OrdinalIgnoreCase))
            {
                // Parse sheet XML and collect elements
                var sheetElements = ExtractSheetElements(entry.Xml, sheetNumber, sheetName, sharedStrings);
                allElements.AddRange(sheetElements);
            }
            // Check if this is a drawing XML file
            else if (entry.PartName.StartsWith("/xl/drawings/drawing", StringComparison.OrdinalIgnoreCase) &&
                     entry.PartName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                // Parse drawing XML and collect elements
                var drawingElements = ExtractDrawingElements(entry.Xml, sheetNumber, sheetName, allElements.Count);
                allElements.AddRange(drawingElements);
            }
            // Skip relationship files - they don't contain useful content for understanding the sheet
        }

        // Write all elements in order
        foreach (var element in allElements)
        {
            writer.WriteLine(JsonSerializer.Serialize(element, JsonContext.SheetElement));
        }
    }

    private static IReadOnlyList<SheetElement> ExtractSheetElements(string xml, int sheetNumber, string sheetName, IReadOnlyList<string> sharedStrings)
    {
        var elements = new List<SheetElement>();
        var elementIndex = 0;

        try
        {
            var doc = XDocument.Parse(xml);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            // Add sheet metadata as first element
            elements.Add(new SheetElement(
                sheetNumber,
                sheetName,
                "sheet_metadata",
                elementIndex++,
                Metadata: new SheetMetadata(sheetNumber, sheetName)
            ));

            // Extract all rows and cells
            var rows = doc.Descendants(ns + "row");
            foreach (var row in rows)
            {
                var rowNum = row.Attribute("r")?.Value;
                int? rowIndex = null;
                if (int.TryParse(rowNum, out var rNum))
                {
                    rowIndex = rNum;
                }

                var cells = row.Descendants(ns + "c");
                foreach (var cell in cells)
                {
                    var cellRef = cell.Attribute("r")?.Value;
                    var cellType = cell.Attribute("t")?.Value; // s = shared string, str = string, b = boolean
                    var cellValue = cell.Element(ns + "v")?.Value;
                    var cellFormula = cell.Element(ns + "f")?.Value;

                    // Parse column from cell reference
                    int? colIndex = null;
                    if (!string.IsNullOrEmpty(cellRef))
                    {
                        colIndex = GetColumnIndex(cellRef);
                    }

                    // Process cell value based on type
                    string? displayValue = cellValue;
                    if (cellType == "s" && !string.IsNullOrEmpty(cellValue))
                    {
                        // Shared string reference
                        if (int.TryParse(cellValue, out var stringIndex) && stringIndex >= 0 && stringIndex < sharedStrings.Count)
                        {
                            displayValue = sharedStrings[stringIndex];
                        }
                    }
                    else if (cellType == "b")
                    {
                        // Boolean
                        displayValue = cellValue == "1" ? "TRUE" : "FALSE";
                    }
                    else if (cellType == "str")
                    {
                        // Inline string
                        var inlineStr = cell.Element(ns + "is")?.Element(ns + "t")?.Value;
                        if (!string.IsNullOrEmpty(inlineStr))
                        {
                            displayValue = inlineStr;
                        }
                    }

                    // Only add cells with content
                    if (!string.IsNullOrEmpty(displayValue) || !string.IsNullOrEmpty(cellFormula))
                    {
                        elements.Add(new SheetElement(
                            sheetNumber,
                            sheetName,
                            "cell",
                            elementIndex++,
                            CellReference: cellRef,
                            Value: displayValue,
                            Formula: cellFormula,
                            DataType: cellType,
                            Row: rowIndex,
                            Column: colIndex
                        ));
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // If XML parsing fails, return sheet as single element with error
            elements.Add(new SheetElement(
                sheetNumber,
                sheetName,
                "raw_xml",
                0,
                ErrorInfo: new ErrorInfo(xml, ex.Message)
            ));
        }

        return elements;
    }

    private static int GetColumnIndex(string cellReference)
    {
        // Extract column letters from cell reference (e.g., "A1" -> "A", "AB12" -> "AB")
        var columnLetters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());

        // Convert column letters to index (A=1, B=2, ..., Z=26, AA=27, etc.)
        int index = 0;
        foreach (char c in columnLetters)
        {
            index = index * 26 + (c - 'A' + 1);
        }
        return index;
    }

    private static IReadOnlyList<SheetElement> ExtractDrawingElements(string xml, int sheetNumber, string sheetName, int startElementIndex)
    {
        var elements = new List<SheetElement>();
        var elementIndex = startElementIndex;

        try
        {
            var doc = XDocument.Parse(xml);
            XNamespace xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            // Process all anchors (two-cell, one-cell, and absolute)
            ProcessAnchors(doc, elements, ref elementIndex, sheetNumber, sheetName, xdr, a);

            return elements;
        }
        catch (Exception ex)
        {
            // If XML parsing fails, return element with error
            elements.Add(new SheetElement(
                sheetNumber,
                sheetName,
                "drawing_error",
                startElementIndex,
                ErrorInfo: new ErrorInfo(xml, ex.Message)
            ));
        }

        return elements;
    }

    private static void ProcessAnchors(XDocument doc, List<SheetElement> elements, ref int elementIndex,
        int sheetNumber, string sheetName, XNamespace xdr, XNamespace a)
    {
        // Extract two-cell anchors (most common for shapes/charts in Excel)
        var twoCellAnchors = doc.Descendants(xdr + "twoCellAnchor");
        foreach (var anchor in twoCellAnchors)
            {
                // Extract anchor positions
                var fromCell = ExtractCellReference(anchor.Element(xdr + "from"), xdr);
                var toCell = ExtractCellReference(anchor.Element(xdr + "to"), xdr);

                CellAnchor? cellAnchor = null;
                if (fromCell != null && toCell != null)
                {
                    cellAnchor = new CellAnchor(
                        fromCell.Value.Cell,
                        fromCell.Value.Col,
                        fromCell.Value.Row,
                        toCell.Value.Cell,
                        toCell.Value.Col,
                        toCell.Value.Row
                    );
                }

                // Process elements within the anchor
                ProcessAnchorElements(anchor, elements, ref elementIndex, sheetNumber, sheetName,
                    xdr, a, cellAnchor, 0, null);
            }

            // Extract absolute position anchors
            var absoluteAnchors = doc.Descendants(xdr + "absoluteAnchor");
            foreach (var anchor in absoluteAnchors)
            {
                var pos = anchor.Element(xdr + "pos");
                var ext = anchor.Element(xdr + "ext");

                Position? position = null;
                Size? size = null;

                if (pos != null)
                {
                    var x = pos.Attribute("x")?.Value;
                    var y = pos.Attribute("y")?.Value;
                    if (long.TryParse(x, out var xVal) && long.TryParse(y, out var yVal))
                    {
                        position = new Position(xVal, yVal);
                    }
                }

                if (ext != null)
                {
                    var cx = ext.Attribute("cx")?.Value;
                    var cy = ext.Attribute("cy")?.Value;
                    if (long.TryParse(cx, out var width) && long.TryParse(cy, out var height))
                    {
                        size = new Size(width, height);
                    }
                }

                var transform = (position != null || size != null) ? new Transform(position, size) : null;

                // Process elements within the anchor
                ProcessAnchorElements(anchor, elements, ref elementIndex, sheetNumber, sheetName,
                    xdr, a, null, 0, null, transform);
            }
        }

    private static void ProcessAnchorElements(XElement anchor, List<SheetElement> elements, ref int elementIndex,
        int sheetNumber, string sheetName, XNamespace xdr, XNamespace a, CellAnchor? cellAnchor,
        int groupLevel, string? parentGroupId, Transform? absoluteTransform = null)
    {
        // Check for group shapes
        var grpSp = anchor.Element(xdr + "grpSp");
        if (grpSp != null)
        {
            ProcessGroupShape(grpSp, elements, ref elementIndex, sheetNumber, sheetName, xdr, a,
                cellAnchor, groupLevel, parentGroupId, absoluteTransform);
            return;
        }

        // Check for shape
        var sp = anchor.Element(xdr + "sp");
        if (sp != null)
        {
            ProcessShape(sp, elements, ref elementIndex, sheetNumber, sheetName, xdr, a,
                cellAnchor, groupLevel, parentGroupId, absoluteTransform);
        }

        // Check for picture
        var pic = anchor.Element(xdr + "pic");
        if (pic != null)
        {
            ProcessPicture(pic, elements, ref elementIndex, sheetNumber, sheetName, xdr, a,
                cellAnchor, groupLevel, parentGroupId, absoluteTransform);
        }

        // Check for graphic frame (charts, tables, SmartArt)
        var graphicFrame = anchor.Element(xdr + "graphicFrame");
        if (graphicFrame != null)
        {
            ProcessGraphicFrame(graphicFrame, elements, ref elementIndex, sheetNumber, sheetName, xdr, a,
                cellAnchor, groupLevel, parentGroupId, absoluteTransform);
        }
    }

    private static void ProcessGroupShape(XElement grpSp, List<SheetElement> elements, ref int elementIndex,
        int sheetNumber, string sheetName, XNamespace xdr, XNamespace a, CellAnchor? cellAnchor,
        int groupLevel, string? parentGroupId, Transform? absoluteTransform)
    {
        var nvGrpSpPr = grpSp.Element(xdr + "nvGrpSpPr");
        var cNvPr = nvGrpSpPr?.Element(xdr + "cNvPr");
        var groupId = cNvPr?.Attribute("id")?.Value;
        var groupName = cNvPr?.Attribute("name")?.Value;

        // Extract group transform
        var grpSpPr = grpSp.Element(xdr + "grpSpPr");
        var xfrm = grpSpPr?.Element(a + "xfrm");
        var groupTransform = xfrm != null ? ExtractTransformFromXfrm(xfrm, a) : absoluteTransform;

        if (groupTransform == null && cellAnchor != null)
        {
            groupTransform = new Transform(Anchor: cellAnchor);
        }

        // Add group element
        elements.Add(new SheetElement(
            sheetNumber,
            sheetName,
            "group",
            elementIndex++,
            Transform: groupTransform,
            ShapeId: groupId,
            ShapeName: groupName,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId
        ));

        // Process child elements
        foreach (var child in grpSp.Elements())
        {
            var childName = child.Name.LocalName;
            if (childName == "sp")
            {
                ProcessShape(child, elements, ref elementIndex, sheetNumber, sheetName, xdr, a,
                    null, groupLevel + 1, groupId, null);
            }
            else if (childName == "grpSp")
            {
                ProcessGroupShape(child, elements, ref elementIndex, sheetNumber, sheetName, xdr, a,
                    null, groupLevel + 1, groupId, null);
            }
            else if (childName == "pic")
            {
                ProcessPicture(child, elements, ref elementIndex, sheetNumber, sheetName, xdr, a,
                    null, groupLevel + 1, groupId, null);
            }
            else if (childName == "graphicFrame")
            {
                ProcessGraphicFrame(child, elements, ref elementIndex, sheetNumber, sheetName, xdr, a,
                    null, groupLevel + 1, groupId, null);
            }
            else if (childName == "cxnSp")
            {
                ProcessConnector(child, elements, ref elementIndex, sheetNumber, sheetName, xdr, a,
                    null, groupLevel + 1, groupId, null);
            }
        }
    }

    private static void ProcessShape(XElement sp, List<SheetElement> elements, ref int elementIndex,
        int sheetNumber, string sheetName, XNamespace xdr, XNamespace a, CellAnchor? cellAnchor,
        int groupLevel, string? parentGroupId, Transform? absoluteTransform)
    {
        var nvSpPr = sp.Element(xdr + "nvSpPr");
        var cNvPr = nvSpPr?.Element(xdr + "cNvPr");
        var shapeId = cNvPr?.Attribute("id")?.Value;
        var shapeName = cNvPr?.Attribute("name")?.Value;

        var spPr = sp.Element(xdr + "spPr");

        // Extract transform
        Transform? transform = null;
        var xfrm = spPr?.Element(a + "xfrm");
        if (xfrm != null)
        {
            transform = ExtractTransformFromXfrm(xfrm, a);
        }
        else if (absoluteTransform != null)
        {
            transform = absoluteTransform;
        }
        else if (cellAnchor != null)
        {
            transform = new Transform(Anchor: cellAnchor);
        }

        // Check for custom geometry
        var custGeom = spPr?.Element(a + "custGeom");
        var customGeometry = XmlUtilities.ExtractCustomGeometry(custGeom, a);

        // Extract shape type
        var prstGeom = spPr?.Element(a + "prstGeom");
        var shapeType = prstGeom?.Attribute("prst")?.Value ?? (customGeometry != null ? "custom" : null);

        elements.Add(new SheetElement(
            sheetNumber,
            sheetName,
            "shape",
            elementIndex++,
            Transform: transform,
            ShapeType: shapeType,
            ShapeId: shapeId,
            ShapeName: shapeName,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId,
            CustomGeometry: customGeometry
        ));
    }

    private static void ProcessPicture(XElement pic, List<SheetElement> elements, ref int elementIndex,
        int sheetNumber, string sheetName, XNamespace xdr, XNamespace a, CellAnchor? cellAnchor,
        int groupLevel, string? parentGroupId, Transform? absoluteTransform)
    {
        var nvPicPr = pic.Element(xdr + "nvPicPr");
        var cNvPr = nvPicPr?.Element(xdr + "cNvPr");
        var picId = cNvPr?.Attribute("id")?.Value;
        var picName = cNvPr?.Attribute("name")?.Value;

        Transform? transform = absoluteTransform;
        if (transform == null && cellAnchor != null)
        {
            transform = new Transform(Anchor: cellAnchor);
        }

        elements.Add(new SheetElement(
            sheetNumber,
            sheetName,
            "image",
            elementIndex++,
            Transform: transform,
            ShapeId: picId,
            ShapeName: picName,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId
        ));
    }

    private static void ProcessGraphicFrame(XElement graphicFrame, List<SheetElement> elements, ref int elementIndex,
        int sheetNumber, string sheetName, XNamespace xdr, XNamespace a, CellAnchor? cellAnchor,
        int groupLevel, string? parentGroupId, Transform? absoluteTransform)
    {
        var nvGraphicFramePr = graphicFrame.Element(xdr + "nvGraphicFramePr");
        var cNvPr = nvGraphicFramePr?.Element(xdr + "cNvPr");
        var frameId = cNvPr?.Attribute("id")?.Value;
        var frameName = cNvPr?.Attribute("name")?.Value;

        Transform? transform = absoluteTransform;
        if (transform == null && cellAnchor != null)
        {
            transform = new Transform(Anchor: cellAnchor);
        }

        // Check graphic data content
        var graphic = graphicFrame.Element(a + "graphic");
        var graphicData = graphic?.Element(a + "graphicData");

        if (graphicData != null)
        {
            // Check for table
            var table = graphicData.Element(a + "tbl");
            if (table != null)
            {
                ProcessTable(table, elements, ref elementIndex, sheetNumber, sheetName, a,
                    frameId, frameName, transform, groupLevel, parentGroupId);
                return;
            }

            // Check for chart
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var chart = graphicData.Element(c + "chart");
            if (chart != null)
            {
                elements.Add(new SheetElement(
                    sheetNumber,
                    sheetName,
                    "chart",
                    elementIndex++,
                    Transform: transform,
                    ShapeId: frameId,
                    ShapeName: frameName,
                    GroupLevel: groupLevel,
                    ParentGroupId: parentGroupId
                ));
                return;
            }

            // Check for SmartArt
            XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            var relIds = graphicData.Element(dgm + "relIds");
            if (relIds != null)
            {
                elements.Add(new SheetElement(
                    sheetNumber,
                    sheetName,
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
                elements.Add(new SheetElement(
                    sheetNumber,
                    sheetName,
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

        // Unknown graphic frame type
        elements.Add(new SheetElement(
            sheetNumber,
            sheetName,
            "graphic_frame",
            elementIndex++,
            Transform: transform,
            ShapeId: frameId,
            ShapeName: frameName,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId
        ));
    }

    private static void ProcessTable(XElement table, List<SheetElement> elements, ref int elementIndex,
        int sheetNumber, string sheetName, XNamespace a, string? tableId, string? tableName,
        Transform? transform, int groupLevel, string? parentGroupId)
    {
        // Add table element
        elements.Add(new SheetElement(
            sheetNumber,
            sheetName,
            "table",
            elementIndex++,
            ShapeId: tableId,
            ShapeName: tableName,
            Transform: transform,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId
        ));

        // Extract table cells using utility
        var cells = XmlUtilities.ExtractTableCells(table, a);
        foreach (var cell in cells)
        {
            if (!string.IsNullOrWhiteSpace(cell.Text))
            {
                elements.Add(new SheetElement(
                    sheetNumber,
                    sheetName,
                    "table_cell",
                    elementIndex++,
                    Value: cell.Text,
                    ShapeId: $"{tableId}_R{cell.Row}C{cell.Col}",
                    ShapeName: $"Cell[{cell.Row},{cell.Col}]",
                    GroupLevel: groupLevel + 1,
                    ParentGroupId: tableId
                ));
            }
        }
    }

    private static void ProcessConnector(XElement cxnSp, List<SheetElement> elements, ref int elementIndex,
        int sheetNumber, string sheetName, XNamespace xdr, XNamespace a, CellAnchor? cellAnchor,
        int groupLevel, string? parentGroupId, Transform? absoluteTransform)
    {
        var nvCxnSpPr = cxnSp.Element(xdr + "nvCxnSpPr");
        var cNvPr = nvCxnSpPr?.Element(xdr + "cNvPr");
        var connectorId = cNvPr?.Attribute("id")?.Value;
        var connectorName = cNvPr?.Attribute("name")?.Value;

        var spPr = cxnSp.Element(xdr + "spPr");

        // Extract transform
        Transform? transform = null;
        var xfrm = spPr?.Element(a + "xfrm");
        if (xfrm != null)
        {
            transform = ExtractTransformFromXfrm(xfrm, a);
        }
        else if (absoluteTransform != null)
        {
            transform = absoluteTransform;
        }
        else if (cellAnchor != null)
        {
            transform = new Transform(Anchor: cellAnchor);
        }

        // Extract connector type
        var prstGeom = spPr?.Element(a + "prstGeom");
        var connectorType = prstGeom?.Attribute("prst")?.Value;

        elements.Add(new SheetElement(
            sheetNumber,
            sheetName,
            "connector",
            elementIndex++,
            Transform: transform,
            ShapeType: connectorType,
            ShapeId: connectorId,
            ShapeName: connectorName,
            GroupLevel: groupLevel,
            ParentGroupId: parentGroupId
        ));
    }


    private static (string Cell, int Col, int Row)? ExtractCellReference(XElement? element, XNamespace xdr)
    {
        return XmlUtilities.ExtractCellReference(element, xdr);
    }

    private static Transform? ExtractTransformFromXfrm(XElement xfrm, XNamespace a)
    {
        return XmlUtilities.ExtractTransformFromXfrm(xfrm, a);
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