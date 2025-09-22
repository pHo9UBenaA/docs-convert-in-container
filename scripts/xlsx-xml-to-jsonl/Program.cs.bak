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
    // Exit codes from CommonBase
    private const int ExitSuccess = CommonBase.ExitSuccess;
    private const int ExitUsageError = CommonBase.ExitUsageError;
    private const int ExitProcessingError = CommonBase.ExitProcessingError;

    // Use custom context with relaxed encoding
    private static readonly SourceGenerationContext JsonContext = SourceGenerationContext.Custom;

    // Type aliases for common types from SharedXmlToJsonl
    private record SheetMetadata(int SheetNumber, string SheetName);
    // Position, Size, Transform, CellAnchor are now from SharedXmlToJsonl namespace

    private record CellValue(
        string? Text = null,
        double? Number = null,
        bool? Boolean = null,
        DateTime? Date = null,
        string ValueType = "text"  // "text", "number", "boolean", "date"
    );

    private record CellFormat(
        int? StyleIndex = null,
        int? NumFmtId = null,
        string? NumFmtCode = null,
        bool? IsDate = null
    );

    private record SheetElement(
        int SheetNumber,
        string SheetName,
        string ElementType,
        int ElementIndex,
        string? CellReference = null,
        CellValue? Value = null,
        string? Formula = null,
        string? DataType = null,
        CellFormat? Format = null,
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
    [JsonSerializable(typeof(CellValue))]
    [JsonSerializable(typeof(CellFormat))]
    [JsonSerializable(typeof(Position))]
    [JsonSerializable(typeof(Size))]
    [JsonSerializable(typeof(CellAnchor))]
    [JsonSerializable(typeof(Transform))]
    [JsonSerializable(typeof(CustomGeometry))]
    [JsonSerializable(typeof(TableCell))]
    private partial class SourceGenerationContext : JsonSerializerContext
    {
        private static readonly JsonSerializerOptions _options = JsonSerializationHelper.CreateStandardOptions();

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
            .Where(PackageUtilities.IsXmlEntry)
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

        // Get cell styles for format lookups
        var cellFormats = ExtractCellFormats(entries);

        foreach (var (sheetNumber, sheetName) in sheetInfos)
        {
            var filePath = Path.Combine(outputDirectory, $"{baseName}_sheet{sheetNumber}.jsonl");
            var sheetSpecificEntries = FilterEntriesForSheet(entries, sheetNumber);
            WriteSheetElementsAsJsonLines(filePath, sheetSpecificEntries, sheetNumber, sheetName, sharedStrings, cellFormats);
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
                var sheetNumber = DocumentUtilities.TryExtractSheetNumber(entry.PartName);
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

    private static IReadOnlyList<CellFormat> ExtractCellFormats(IReadOnlyList<JsonlEntry> entries)
    {
        var cellFormats = new List<CellFormat>();
        var stylesEntry = entries.FirstOrDefault(e => e.PartName == "/xl/styles.xml");

        if (stylesEntry != null)
        {
            try
            {
                var doc = XDocument.Parse(stylesEntry.Xml);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                // Extract number formats (both built-in and custom)
                var customFormats = new Dictionary<int, string>();
                var numFmts = doc.Descendants(ns + "numFmts").FirstOrDefault();
                if (numFmts != null)
                {
                    foreach (var numFmt in numFmts.Elements(ns + "numFmt"))
                    {
                        var numFmtId = numFmt.Attribute("numFmtId")?.Value;
                        var formatCode = numFmt.Attribute("formatCode")?.Value;
                        if (int.TryParse(numFmtId, out var id) && formatCode != null)
                        {
                            customFormats[id] = formatCode;
                        }
                    }
                }

                // Extract cell formats (xf elements in cellXfs)
                var cellXfs = doc.Descendants(ns + "cellXfs").FirstOrDefault();
                if (cellXfs != null)
                {
                    int styleIndex = 0;
                    foreach (var xf in cellXfs.Elements(ns + "xf"))
                    {
                        var numFmtIdAttr = xf.Attribute("numFmtId")?.Value;
                        int? numFmtId = null;
                        if (int.TryParse(numFmtIdAttr, out var fmtId))
                        {
                            numFmtId = fmtId;
                        }

                        string? numFmtCode = null;
                        bool isDate = false;

                        if (numFmtId.HasValue)
                        {
                            // Check if it's a custom format
                            if (customFormats.ContainsKey(numFmtId.Value))
                            {
                                numFmtCode = customFormats[numFmtId.Value];
                            }
                            // Check for built-in date formats
                            // Date format IDs: 14-22, 27-36, 45-47, 50-58, 71-81
                            isDate = IsDateFormatId(numFmtId.Value);

                            // Also check if custom format contains date patterns
                            if (!isDate && numFmtCode != null)
                            {
                                isDate = IsDateFormatCode(numFmtCode);
                            }
                        }

                        cellFormats.Add(new CellFormat(
                            StyleIndex: styleIndex,
                            NumFmtId: numFmtId,
                            NumFmtCode: numFmtCode,
                            IsDate: isDate
                        ));

                        styleIndex++;
                    }
                }
            }
            catch
            {
                // If parsing fails, return empty list
            }
        }

        return cellFormats;
    }

    private static bool IsDateFormatId(int numFmtId)
    {
        // Standard Excel date/time format IDs
        return (numFmtId >= 14 && numFmtId <= 22) ||
               (numFmtId >= 27 && numFmtId <= 36) ||
               (numFmtId >= 45 && numFmtId <= 47) ||
               (numFmtId >= 50 && numFmtId <= 58) ||
               (numFmtId >= 71 && numFmtId <= 81);
    }

    private static bool IsDateFormatCode(string formatCode)
    {
        // Check if format code contains date/time patterns
        var datePatterns = new[] { "yy", "mm", "dd", "hh", "ss", "AM/PM", "A/P", "[h]", "[m]", "[s]" };
        var lowerCode = formatCode.ToLowerInvariant();
        return datePatterns.Any(pattern => lowerCode.Contains(pattern.ToLowerInvariant()));
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
        using var reader = new StreamReader(stream, CommonBase.Utf8NoBom, detectEncodingFromByteOrderMarks: true);
        var xml = reader.ReadToEnd();
        var sizeBytes = CommonBase.Utf8NoBom.GetByteCount(xml);

        if (PackageUtilities.TryGetPackagePart(package, partName, out var packagePart))
        {
            var relationships = PackageUtilities.ExtractRelationships(packagePart);
            return new JsonlEntry(partName, packagePart.ContentType, relationships, sizeBytes, xml);
        }

        var fallbackContentType = PackageUtilities.DetermineFallbackContentType(partName);
        return new JsonlEntry(partName, fallbackContentType, Array.Empty<RelationshipInfo>(), sizeBytes, xml);
    }



    private static void WriteSheetElementsAsJsonLines(string outputPath, IReadOnlyList<JsonlEntry> entries, int sheetNumber, string sheetName, IReadOnlyList<string> sharedStrings, IReadOnlyList<CellFormat> cellFormats)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var stream = File.Open(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, CommonBase.Utf8NoBom);

        var allElements = new List<SheetElement>();

        foreach (var entry in entries)
        {
            // Check if this is a sheet XML file
            if (entry.PartName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
                entry.PartName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !entry.PartName.Contains(".rels", StringComparison.OrdinalIgnoreCase))
            {
                // Parse sheet XML and collect elements
                var sheetElements = ExtractSheetElements(entry.Xml, sheetNumber, sheetName, sharedStrings, cellFormats);
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

    private static IReadOnlyList<SheetElement> ExtractSheetElements(string xml, int sheetNumber, string sheetName, IReadOnlyList<string> sharedStrings, IReadOnlyList<CellFormat> cellFormats)
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
                    var cellStyleAttr = cell.Attribute("s")?.Value; // style index
                    var cellValue = cell.Element(ns + "v")?.Value;
                    var cellFormula = cell.Element(ns + "f")?.Value;

                    // Parse column from cell reference
                    int? colIndex = null;
                    if (!string.IsNullOrEmpty(cellRef))
                    {
                        colIndex = GetColumnIndex(cellRef);
                    }

                    // Get cell format information
                    CellFormat? cellFormat = null;
                    if (int.TryParse(cellStyleAttr, out var styleIndex) && styleIndex >= 0 && styleIndex < cellFormats.Count)
                    {
                        cellFormat = cellFormats[styleIndex];
                    }

                    // Process cell value based on type and format
                    CellValue? processedValue = null;
                    if (!string.IsNullOrEmpty(cellValue) || cellType == "str")
                    {
                        processedValue = ProcessCellValue(cellValue, cellType, cellFormat, sharedStrings, cell, ns);
                    }

                    // Only add cells with content
                    if (processedValue != null || !string.IsNullOrEmpty(cellFormula))
                    {
                        elements.Add(new SheetElement(
                            sheetNumber,
                            sheetName,
                            "cell",
                            elementIndex++,
                            CellReference: cellRef,
                            Value: processedValue,
                            Formula: cellFormula,
                            DataType: cellType,
                            Format: cellFormat,
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

    private static CellValue? ProcessCellValue(string? cellValue, string? cellType, CellFormat? cellFormat, IReadOnlyList<string> sharedStrings, XElement cell, XNamespace ns)
    {
        if (cellType == "s" && !string.IsNullOrEmpty(cellValue))
        {
            // Shared string reference
            if (int.TryParse(cellValue, out var stringIndex) && stringIndex >= 0 && stringIndex < sharedStrings.Count)
            {
                return new CellValue(Text: sharedStrings[stringIndex], ValueType: "text");
            }
        }
        else if (cellType == "b" && !string.IsNullOrEmpty(cellValue))
        {
            // Boolean
            var boolValue = cellValue == "1";
            return new CellValue(Boolean: boolValue, Text: boolValue ? "TRUE" : "FALSE", ValueType: "boolean");
        }
        else if (cellType == "str")
        {
            // Inline string
            var inlineStr = cell.Element(ns + "is")?.Element(ns + "t")?.Value;
            if (!string.IsNullOrEmpty(inlineStr))
            {
                return new CellValue(Text: inlineStr, ValueType: "text");
            }
        }
        else if (!string.IsNullOrEmpty(cellValue))
        {
            // Numeric value (could be number or date)
            if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var numValue))
            {
                // Check if it's a date based on format
                if (cellFormat?.IsDate == true)
                {
                    // Convert Excel date serial number to DateTime
                    var dateTime = ConvertExcelSerialToDateTime(numValue);
                    if (dateTime.HasValue)
                    {
                        return new CellValue(
                            Date: dateTime.Value,
                            Text: dateTime.Value.ToString("yyyy-MM-dd'T'HH:mm:ss"),
                            ValueType: "date"
                        );
                    }
                }

                // Return as number
                return new CellValue(
                    Number: numValue,
                    Text: numValue.ToString(CultureInfo.InvariantCulture),
                    ValueType: "number"
                );
            }
            else
            {
                // If parsing as number fails, treat as text
                return new CellValue(Text: cellValue, ValueType: "text");
            }
        }

        return null;
    }

    private static DateTime? ConvertExcelSerialToDateTime(double serialDate)
    {
        // Excel dates start from 1900-01-01 (serial number 1)
        // But Excel incorrectly treats 1900 as a leap year, so we need to handle this
        const int excelBaseYear = 1900;
        const int dayAdjustment = -2; // Adjustment for Excel's 1900 leap year bug

        if (serialDate < 1)
        {
            return null; // Invalid date
        }

        try
        {
            if (serialDate <= 60)
            {
                // Before March 1, 1900 (no leap year bug adjustment needed)
                return new DateTime(excelBaseYear, 1, 1).AddDays(serialDate - 1);
            }
            else
            {
                // After February 28, 1900 (apply leap year bug adjustment)
                return new DateTime(excelBaseYear, 1, 1).AddDays(serialDate + dayAdjustment);
            }
        }
        catch
        {
            return null;
        }
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
            XNamespace xdr = NamespaceConstants.XDR;
            XNamespace a = NamespaceConstants.A;

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
        var groupTransform = xfrm != null ? XmlUtilities.ExtractTransformFromXfrm(xfrm, a) : absoluteTransform;

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
            transform = XmlUtilities.ExtractTransformFromXfrm(xfrm, a);
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
        var shapeType = ShapeProcessor.ExtractShapeType(spPr, a);

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
                    Value: new CellValue(Text: cell.Text, ValueType: "text"),
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
        var currentIndex = elementIndex;
        ShapeProcessor.ProcessConnectorXlsx(
            cxnSp,
            elements,
            ref elementIndex,
            xdr,
            a,
            (connId, connName, transform, shapeType, grpLevel, parentGrpId) =>
                new SheetElement(
                    sheetNumber,
                    sheetName,
                    "connector",
                    currentIndex,
                    Transform: transform,
                    ShapeType: shapeType,
                    ShapeId: connId,
                    ShapeName: connName,
                    GroupLevel: grpLevel,
                    ParentGroupId: parentGrpId
                ),
            absoluteTransform,
            cellAnchor,
            groupLevel,
            parentGroupId
        );
    }


    private static (string Cell, int Col, int Row)? ExtractCellReference(XElement? element, XNamespace xdr)
    {
        return XmlUtilities.ExtractCellReference(element, xdr);
    }


    private static void WriteJsonLines(string outputPath, IReadOnlyList<JsonlEntry> entries)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var stream = File.Open(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, CommonBase.Utf8NoBom);

        foreach (var entry in entries)
        {
            writer.WriteLine(JsonSerializer.Serialize(entry, JsonContext.JsonlEntry));
        }
    }
}