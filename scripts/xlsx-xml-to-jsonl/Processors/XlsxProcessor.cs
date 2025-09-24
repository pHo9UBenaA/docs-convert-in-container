using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Extensions.Logging;
using SharedXmlToJsonl;
using SharedXmlToJsonl.Configuration;
using SharedXmlToJsonl.Factories;
using SharedXmlToJsonl.Interfaces;
using SharedXmlToJsonl.Models;

namespace XlsxXmlToJsonl.Processors;

public class XlsxProcessor : IXlsxProcessor
{
    private readonly IElementFactory? _elementFactory;
    private readonly IJsonWriter? _jsonWriter;
    private readonly ILogger<XlsxProcessor>? _logger;

    public XlsxProcessor(
        IElementFactory? elementFactory,
        IJsonWriter? jsonWriter,
        ILogger<XlsxProcessor>? logger)
    {
        _elementFactory = elementFactory;
        _jsonWriter = jsonWriter;
        _logger = logger;
    }

    // Implementation of IDocumentProcessor.ProcessAsync
    public async Task<ProcessingResult> ProcessAsync(
        string inputPath,
        string outputDirectory,
        CancellationToken cancellationToken = default)
    {
        // Use default options when called from IDocumentProcessor interface
        var options = new ProcessingOptions();
        return await ProcessAsync(inputPath, outputDirectory, options, cancellationToken);
    }

    // Implementation of IXlsxProcessor.ProcessAsync
    public async Task<ProcessingResult> ProcessAsync(
        string inputPath,
        string outputDirectory,
        ProcessingOptions options,
        CancellationToken cancellationToken = default)
    {
        _logger?.LogInformation("Starting XLSX processing for {Path}", inputPath);

        try
        {
            var sheetDataByNumber = await ExtractSheetsAsync(inputPath, cancellationToken);

            if (!sheetDataByNumber.Any())
            {
                _logger?.LogWarning("No worksheets found in {Path}", inputPath);
                return new ProcessingResult
                {
                    Success = false,
                    ErrorMessage = "No worksheets found in the workbook"
                };
            }

            var baseName = Path.GetFileNameWithoutExtension(inputPath);
            var outputPaths = new List<string>();

            // Write each sheet to a separate file
            foreach (var (sheetNumber, sheetElements) in sheetDataByNumber)
            {
                // Get sheet name from the first element (sheet_metadata)
                var sheetName = sheetElements.FirstOrDefault()?.SheetName ?? $"Sheet{sheetNumber}";
                var sanitizedSheetName = SanitizeFileName(sheetName);
                var outputPath = Path.Combine(outputDirectory, $"sheet-{sanitizedSheetName}.jsonl");
                await WriteSheetToFile(outputPath, sheetElements, cancellationToken);
                outputPaths.Add(outputPath);
            }

            _logger?.LogInformation("Successfully processed {Count} sheets", sheetDataByNumber.Count);
            return new ProcessingResult
            {
                Success = true,
                ItemsProcessed = sheetDataByNumber.Count,
                OutputPath = string.Join(";", outputPaths)
            };
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error processing XLSX file {Path}", inputPath);
            return new ProcessingResult
            {
                Success = false,
                ErrorMessage = ex.Message
            };
        }
    }

    private async Task WriteSheetToFile(
        string outputPath,
        List<SheetElement> sheetElements,
        CancellationToken cancellationToken)
    {
        await using var fileStream = new FileStream(
            outputPath,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 4096,
            useAsync: true);
        await using var writer = new StreamWriter(fileStream);

        var options = new JsonSerializerOptions(ElementJsonSerializerContext.Default.SheetElement.Options)
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        foreach (var element in sheetElements)
        {
            var json = JsonSerializer.Serialize(element, options);
            await writer.WriteLineAsync(json);
        }
    }

    private async Task<Dictionary<int, List<SheetElement>>> ExtractSheetsAsync(
        string path,
        CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            using var package = Package.Open(path, FileMode.Open, FileAccess.Read);
            var sheetDataByNumber = new Dictionary<int, List<SheetElement>>();

            // Load shared strings
            var sharedStrings = LoadSharedStrings(package);

            // Load styles for date detection
            var dateFormats = LoadDateFormats(package);

            // Get workbook part
            var workbookPart = PackageUtilities.GetWorkbookPart(package);
            if (workbookPart == null)
            {
                _logger?.LogWarning("No workbook part found in {Path}", path);
                return sheetDataByNumber;
            }

            var workbookDoc = PackageUtilities.GetXDocument(workbookPart);
            var sheets = workbookDoc
                .Descendants(NamespaceConstants.spreadsheet + "sheets")
                .Elements(NamespaceConstants.spreadsheet + "sheet")
                .ToList();

            var sheetNumber = 1;
            foreach (var sheet in sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var sheetName = sheet.Attribute("name")?.Value ?? $"Sheet{sheetNumber}";
                var rId = sheet.Attribute(NamespaceConstants.R + "id")?.Value;

                if (string.IsNullOrEmpty(rId))
                    continue;

                var sheetRelationship = workbookPart.GetRelationship(rId);
                var worksheetPart = package.GetPart(
                    PackUriHelper.ResolvePartUri(workbookPart.Uri, sheetRelationship.TargetUri));

                var worksheetDoc = PackageUtilities.GetXDocument(worksheetPart);
                var sheetElements = ProcessWorksheet(worksheetDoc, sharedStrings, dateFormats, sheetNumber, sheetName);

                // Process drawings
                var elementIndex = sheetElements.Count;
                var drawingElements = ProcessDrawings(package, worksheetPart, sheetNumber, sheetName, ref elementIndex);
                sheetElements.AddRange(drawingElements);

                sheetDataByNumber[sheetNumber] = sheetElements;
                sheetNumber++;
            }

            return sheetDataByNumber;
        }, cancellationToken);
    }

    private Dictionary<int, string> LoadSharedStrings(Package package)
    {
        var sharedStringsDict = new Dictionary<int, string>();
        var sharedStringsPart = PackageUtilities.GetSharedStringsPart(package);

        if (sharedStringsPart != null)
        {
            var sharedStringsDoc = PackageUtilities.GetXDocument(sharedStringsPart);
            var stringItems = sharedStringsDoc
                .Descendants(NamespaceConstants.spreadsheet + "si")
                .ToList();

            for (int i = 0; i < stringItems.Count; i++)
            {
                var text = ExtractTextFromSi(stringItems[i]);
                sharedStringsDict[i] = text;
            }
        }

        return sharedStringsDict;
    }

    private HashSet<int> LoadDateFormats(Package package)
    {
        var dateFormats = new HashSet<int>();
        var stylesPart = PackageUtilities.GetStylesPart(package);

        if (stylesPart != null)
        {
            var stylesDoc = PackageUtilities.GetXDocument(stylesPart);

            // Get custom number formats that look like dates
            var numFmts = stylesDoc
                .Descendants(NamespaceConstants.spreadsheet + "numFmt")
                .Where(n => IsDateFormat(n.Attribute("formatCode")?.Value ?? ""))
                .Select(n => int.Parse(n.Attribute("numFmtId")?.Value ?? "0"));

            foreach (var numFmtId in numFmts)
            {
                dateFormats.Add(numFmtId);
            }

            // Add built-in date format IDs
            for (int i = 14; i <= 22; i++)
                dateFormats.Add(i);
            for (int i = 45; i <= 47; i++)
                dateFormats.Add(i);
        }

        return dateFormats;
    }

    private bool IsDateFormat(string formatCode)
    {
        var dateIndicators = new[] { "mm", "dd", "yy", "hh", "ss", "AM/PM" };
        return dateIndicators.Any(indicator => formatCode.Contains(indicator, StringComparison.OrdinalIgnoreCase));
    }

    private string ExtractTextFromSi(XElement si)
    {
        var texts = new List<string>();

        // Check for simple text
        var t = si.Element(NamespaceConstants.spreadsheet + "t");
        if (t != null)
        {
            return t.Value;
        }

        // Check for rich text
        foreach (var r in si.Elements(NamespaceConstants.spreadsheet + "r"))
        {
            var rText = r.Element(NamespaceConstants.spreadsheet + "t")?.Value;
            if (!string.IsNullOrEmpty(rText))
                texts.Add(rText);
        }

        return string.Join("", texts);
    }

    private List<SheetElement> ProcessWorksheet(
        XDocument worksheetDoc,
        Dictionary<int, string> sharedStrings,
        HashSet<int> dateFormats,
        int sheetNumber,
        string sheetName)
    {
        var elements = new List<SheetElement>();
        var elementIndex = 0;

        // Load merge cells information
        var mergeCells = LoadMergeCells(worksheetDoc);

        // Add sheet metadata as first element
        elements.Add(new SheetElement
        {
            SheetNumber = sheetNumber,
            SheetName = sheetName,
            ElementType = "sheet_metadata",
            ElementIndex = elementIndex++,
            Metadata = new Dictionary<string, object>
            {
                ["sheet_number"] = sheetNumber,
                ["sheet_name"] = sheetName
            }
        });

        var sheetData = worksheetDoc.Root?.Element(NamespaceConstants.spreadsheet + "sheetData");
        if (sheetData == null)
            return elements;

        // Process all cells
        foreach (var row in sheetData.Elements(NamespaceConstants.spreadsheet + "row"))
        {
            var rowNum = int.Parse(row.Attribute("r")?.Value ?? "0");

            foreach (var cell in row.Elements(NamespaceConstants.spreadsheet + "c"))
            {
                var cellElement = ProcessCell(cell, sharedStrings, dateFormats, mergeCells, sheetNumber, sheetName, ref elementIndex);
                if (cellElement != null)
                {
                    cellElement.Row = rowNum;
                    elements.Add(cellElement);
                }
            }
        }

        return elements;
    }

    /// <summary>
    /// Process drawings (shapes and images) in the worksheet
    /// </summary>
    private List<SheetElement> ProcessDrawings(
        Package package,
        PackagePart worksheetPart,
        int sheetNumber,
        string sheetName,
        ref int elementIndex)
    {
        var elements = new List<SheetElement>();

        try
        {
            // Find drawing relationship
            var drawingRel = worksheetPart.GetRelationships()
                .FirstOrDefault(r => r.RelationshipType == NamespaceConstants.DrawingRelType);

            if (drawingRel == null)
                return elements;

            // Get drawing part
            var drawingPart = package.GetPart(
                PackUriHelper.ResolvePartUri(worksheetPart.Uri, drawingRel.TargetUri));

            var drawingDoc = PackageUtilities.GetXDocument(drawingPart);

            // Process two-cell anchors (most common for shapes/images)
            foreach (var twoCellAnchor in drawingDoc.Descendants(NamespaceConstants.XDR + "twoCellAnchor"))
            {
                var anchorInfo = ExtractTwoCellAnchor(twoCellAnchor);

                // Process picture
                var pic = twoCellAnchor.Element(NamespaceConstants.XDR + "pic");
                if (pic != null)
                {
                    var pictureElement = ExtractPictureElement(pic, drawingPart, package, sheetNumber, sheetName, ref elementIndex);
                    if (pictureElement != null)
                    {
                        pictureElement.AnchorType = "twoCellAnchor";
                        pictureElement.AnchorFrom = anchorInfo.from;
                        pictureElement.AnchorTo = anchorInfo.to;
                        elements.Add(pictureElement);
                    }
                }

                // Process shape
                var sp = twoCellAnchor.Element(NamespaceConstants.XDR + "sp");
                if (sp != null)
                {
                    var shapeElement = ExtractShapeElement(sp, sheetNumber, sheetName, ref elementIndex);
                    if (shapeElement != null)
                    {
                        shapeElement.AnchorType = "twoCellAnchor";
                        shapeElement.AnchorFrom = anchorInfo.from;
                        shapeElement.AnchorTo = anchorInfo.to;
                        elements.Add(shapeElement);
                    }
                }

                // Process graphic frame (charts, etc.)
                var graphicFrame = twoCellAnchor.Element(NamespaceConstants.XDR + "graphicFrame");
                if (graphicFrame != null)
                {
                    var graphicElement = ExtractGraphicFrameElement(graphicFrame, sheetNumber, sheetName, ref elementIndex);
                    if (graphicElement != null)
                    {
                        graphicElement.AnchorType = "twoCellAnchor";
                        graphicElement.AnchorFrom = anchorInfo.from;
                        graphicElement.AnchorTo = anchorInfo.to;
                        elements.Add(graphicElement);
                    }
                }

                // Process connector
                var cxnSp = twoCellAnchor.Element(NamespaceConstants.XDR + "cxnSp");
                if (cxnSp != null)
                {
                    var connectorElement = ExtractConnectorElement(cxnSp, sheetNumber, sheetName, ref elementIndex);
                    if (connectorElement != null)
                    {
                        connectorElement.AnchorType = "twoCellAnchor";
                        connectorElement.AnchorFrom = anchorInfo.from;
                        connectorElement.AnchorTo = anchorInfo.to;
                        elements.Add(connectorElement);
                    }
                }
            }

            // Process one-cell anchors
            foreach (var oneCellAnchor in drawingDoc.Descendants(NamespaceConstants.XDR + "oneCellAnchor"))
            {
                var anchorInfo = ExtractOneCellAnchor(oneCellAnchor);

                // Process picture elements
                var pic = oneCellAnchor.Element(NamespaceConstants.XDR + "pic");
                if (pic != null)
                {
                    var pictureElement = ExtractPictureElement(pic, drawingPart, package, sheetNumber, sheetName, ref elementIndex);
                    if (pictureElement != null)
                    {
                        pictureElement.AnchorType = "oneCellAnchor";
                        pictureElement.AnchorFrom = anchorInfo;
                        elements.Add(pictureElement);
                    }
                }

                // Process shape elements
                var sp = oneCellAnchor.Element(NamespaceConstants.XDR + "sp");
                if (sp != null)
                {
                    var shapeElement = ExtractShapeElement(sp, sheetNumber, sheetName, ref elementIndex);
                    if (shapeElement != null)
                    {
                        shapeElement.AnchorType = "oneCellAnchor";
                        shapeElement.AnchorFrom = anchorInfo;
                        elements.Add(shapeElement);
                    }
                }

                // Process group shape elements
                var grpSp = oneCellAnchor.Element(NamespaceConstants.XDR + "grpSp");
                if (grpSp != null)
                {
                    // Process shapes within the group
                    foreach (var groupedSp in grpSp.Elements(NamespaceConstants.XDR + "sp"))
                    {
                        var groupedShapeElement = ExtractShapeElement(groupedSp, sheetNumber, sheetName, ref elementIndex);
                        if (groupedShapeElement != null)
                        {
                            groupedShapeElement.AnchorType = "oneCellAnchor";
                            groupedShapeElement.AnchorFrom = anchorInfo;
                            elements.Add(groupedShapeElement);
                        }
                    }

                    // Process connectors within the group
                    foreach (var cxnSp in grpSp.Elements(NamespaceConstants.XDR + "cxnSp"))
                    {
                        var connectorElement = ExtractConnectorElement(cxnSp, sheetNumber, sheetName, ref elementIndex);
                        if (connectorElement != null)
                        {
                            connectorElement.AnchorType = "oneCellAnchor";
                            connectorElement.AnchorFrom = anchorInfo;
                            elements.Add(connectorElement);
                        }
                    }
                }
            }

            // Process absolute anchors
            foreach (var absoluteAnchor in drawingDoc.Descendants(NamespaceConstants.XDR + "absoluteAnchor"))
            {
                // Process elements with absolute positioning
                var pic = absoluteAnchor.Element(NamespaceConstants.XDR + "pic");
                if (pic != null)
                {
                    var pictureElement = ExtractPictureElement(pic, drawingPart, package, sheetNumber, sheetName, ref elementIndex);
                    if (pictureElement != null)
                    {
                        pictureElement.AnchorType = "absoluteAnchor";
                        elements.Add(pictureElement);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Error processing drawings for sheet {SheetNumber}", sheetNumber);
        }

        return elements;
    }

    /// <summary>
    /// Extract two-cell anchor information
    /// </summary>
    private (CellAnchor from, CellAnchor to) ExtractTwoCellAnchor(XElement twoCellAnchor)
    {
        var from = twoCellAnchor.Element(NamespaceConstants.XDR + "from");
        var to = twoCellAnchor.Element(NamespaceConstants.XDR + "to");

        var fromAnchor = ExtractCellAnchor(from);
        var toAnchor = ExtractCellAnchor(to);

        return (fromAnchor, toAnchor);
    }

    /// <summary>
    /// Extract one-cell anchor information
    /// </summary>
    private CellAnchor ExtractOneCellAnchor(XElement oneCellAnchor)
    {
        var from = oneCellAnchor.Element(NamespaceConstants.XDR + "from");
        return ExtractCellAnchor(from);
    }

    /// <summary>
    /// Extract cell anchor details
    /// </summary>
    private CellAnchor ExtractCellAnchor(XElement? anchorElement)
    {
        if (anchorElement == null)
            return new CellAnchor("", 0, 0, "", 0, 0);

        var col = int.Parse(anchorElement.Element(NamespaceConstants.XDR + "col")?.Value ?? "0");
        var colOff = int.Parse(anchorElement.Element(NamespaceConstants.XDR + "colOff")?.Value ?? "0");
        var row = int.Parse(anchorElement.Element(NamespaceConstants.XDR + "row")?.Value ?? "0");
        var rowOff = int.Parse(anchorElement.Element(NamespaceConstants.XDR + "rowOff")?.Value ?? "0");

        var cellRef = GetCellReference(col + 1, row + 1); // Convert to 1-based

        return new CellAnchor(cellRef, col, row, cellRef, col, row);
    }

    /// <summary>
    /// Extract picture element
    /// </summary>
    private SheetElement? ExtractPictureElement(
        XElement pic,
        PackagePart drawingPart,
        Package package,
        int sheetNumber,
        string sheetName,
        ref int elementIndex)
    {
        var nvPicPr = pic.Element(NamespaceConstants.XDR + "nvPicPr");
        var cNvPr = nvPicPr?.Element(NamespaceConstants.XDR + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";
        var descr = cNvPr?.Attribute("descr")?.Value;

        // Get image reference
        var blipFill = pic.Element(NamespaceConstants.XDR + "blipFill");
        var blip = blipFill?.Element(NamespaceConstants.A + "blip");
        var embedId = blip?.Attribute(NamespaceConstants.R + "embed")?.Value;

        string? imagePath = null;
        if (!string.IsNullOrEmpty(embedId))
        {
            imagePath = ResolveImagePath(drawingPart, embedId, package);
        }

        // Get shape properties
        var spPr = pic.Element(NamespaceConstants.XDR + "spPr");
        var transform = ExtractTransformFromSpPr(spPr);

        return new SheetElement
        {
            SheetNumber = sheetNumber,
            SheetName = sheetName,
            ElementType = "picture",
            ElementIndex = elementIndex++,
            ShapeId = id,
            ShapeName = name,
            ShapeType = "picture",
            Transform = transform,
            ImagePath = imagePath,
            Metadata = descr != null ? new Dictionary<string, object> { ["description"] = descr } : null
        };
    }

    /// <summary>
    /// Extract shape element
    /// </summary>
    private SheetElement? ExtractShapeElement(
        XElement sp,
        int sheetNumber,
        string sheetName,
        ref int elementIndex)
    {
        var nvSpPr = sp.Element(NamespaceConstants.XDR + "nvSpPr");
        var cNvPr = nvSpPr?.Element(NamespaceConstants.XDR + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";

        var spPr = sp.Element(NamespaceConstants.XDR + "spPr");
        var transform = ExtractTransformFromSpPr(spPr);

        // Get shape type
        var prstGeom = spPr?.Element(NamespaceConstants.A + "prstGeom");
        var shapeType = prstGeom?.Attribute("prst")?.Value;

        // Extract text from txBody element
        var txBody = sp.Element(NamespaceConstants.XDR + "txBody");
        var text = ExtractTextFromTxBody(txBody);

        var sheetElement = new SheetElement
        {
            SheetNumber = sheetNumber,
            SheetName = sheetName,
            ElementType = "shape",
            ElementIndex = elementIndex++,
            ShapeId = id,
            ShapeName = name,
            ShapeType = shapeType,
            Transform = transform
        };

        // Set Value if text is present
        if (!string.IsNullOrEmpty(text))
        {
            sheetElement.Value = new CellValue
            {
                Text = text,
                ValueType = "text"
            };
        }

        return sheetElement;
    }

    /// <summary>
    /// Extract graphic frame element (charts, etc.)
    /// </summary>
    private SheetElement? ExtractGraphicFrameElement(
        XElement graphicFrame,
        int sheetNumber,
        string sheetName,
        ref int elementIndex)
    {
        var nvGraphicFramePr = graphicFrame.Element(NamespaceConstants.XDR + "nvGraphicFramePr");
        var cNvPr = nvGraphicFramePr?.Element(NamespaceConstants.XDR + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";

        var xfrm = graphicFrame.Element(NamespaceConstants.XDR + "xfrm");
        var transform = ExtractTransformFromXfrm(xfrm);

        // Determine the type of graphic
        var graphic = graphicFrame.Element(NamespaceConstants.A + "graphic");
        var graphicData = graphic?.Element(NamespaceConstants.A + "graphicData");
        var uri = graphicData?.Attribute("uri")?.Value;

        string? graphicType = "graphic";
        if (uri != null && uri.Contains("chart"))
        {
            graphicType = "chart";
        }
        else if (uri != null && uri.Contains("diagram"))
        {
            graphicType = "diagram";
        }

        return new SheetElement
        {
            SheetNumber = sheetNumber,
            SheetName = sheetName,
            ElementType = "graphic_frame",
            ElementIndex = elementIndex++,
            ShapeId = id,
            ShapeName = name,
            ShapeType = graphicType,
            Transform = transform
        };
    }

    /// <summary>
    /// Extract connector element
    /// </summary>
    private SheetElement? ExtractConnectorElement(
        XElement cxnSp,
        int sheetNumber,
        string sheetName,
        ref int elementIndex)
    {
        var nvCxnSpPr = cxnSp.Element(NamespaceConstants.XDR + "nvCxnSpPr");
        var cNvPr = nvCxnSpPr?.Element(NamespaceConstants.XDR + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";

        // Extract connection information
        var cNvCxnSpPr = nvCxnSpPr?.Element(NamespaceConstants.XDR + "cNvCxnSpPr");
        ConnectionInfo? startConnection = null;
        ConnectionInfo? endConnection = null;

        if (cNvCxnSpPr != null)
        {
            // Extract start connection
            var stCxn = cNvCxnSpPr.Element(NamespaceConstants.A + "stCxn");
            if (stCxn != null)
            {
                var startId = stCxn.Attribute("id")?.Value;
                var startIdx = stCxn.Attribute("idx")?.Value;
                if (startId != null)
                {
                    startConnection = new ConnectionInfo(
                        ShapeId: startId,
                        ConnectionSiteIndex: int.TryParse(startIdx, out var idx) ? idx : null
                    );
                }
            }

            // Extract end connection
            var endCxn = cNvCxnSpPr.Element(NamespaceConstants.A + "endCxn");
            if (endCxn != null)
            {
                var endId = endCxn.Attribute("id")?.Value;
                var endIdx = endCxn.Attribute("idx")?.Value;
                if (endId != null)
                {
                    endConnection = new ConnectionInfo(
                        ShapeId: endId,
                        ConnectionSiteIndex: int.TryParse(endIdx, out var idx) ? idx : null
                    );
                }
            }
        }

        var spPr = cxnSp.Element(NamespaceConstants.XDR + "spPr");
        var transform = ExtractTransformFromSpPr(spPr);

        // Extract text from txBody element
        var txBody = cxnSp.Element(NamespaceConstants.XDR + "txBody");
        var text = ExtractTextFromTxBody(txBody);

        var sheetElement = new SheetElement
        {
            SheetNumber = sheetNumber,
            SheetName = sheetName,
            ElementType = "connector",
            ElementIndex = elementIndex++,
            ShapeId = id,
            ShapeName = name,
            ShapeType = "connector",
            Transform = transform,
            StartConnection = startConnection,
            EndConnection = endConnection
        };

        // Set Value if text is present
        if (!string.IsNullOrEmpty(text))
        {
            sheetElement.Value = new CellValue
            {
                Text = text,
                ValueType = "text"
            };
        }

        return sheetElement;
    }

    /// <summary>
    /// Extract transform from shape properties
    /// </summary>
    private SharedXmlToJsonl.Models.Transform? ExtractTransformFromSpPr(XElement? spPr)
    {
        if (spPr == null)
            return null;

        var xfrm = spPr.Element(NamespaceConstants.A + "xfrm");
        return ExtractTransformFromXfrm(xfrm);
    }

    /// <summary>
    /// Extract transform from xfrm element
    /// </summary>
    private SharedXmlToJsonl.Models.Transform? ExtractTransformFromXfrm(XElement? xfrm)
    {
        if (xfrm == null)
            return null;

        var transform = new SharedXmlToJsonl.Models.Transform();

        var off = xfrm.Element(NamespaceConstants.A + "off");
        if (off != null)
        {
            var x = off.Attribute("x")?.Value;
            var y = off.Attribute("y")?.Value;
            if (x != null && y != null && long.TryParse(x, out long xVal) && long.TryParse(y, out long yVal))
            {
                transform.Position = new Position(xVal, yVal);
            }
        }

        var ext = xfrm.Element(NamespaceConstants.A + "ext");
        if (ext != null)
        {
            var cx = ext.Attribute("cx")?.Value;
            var cy = ext.Attribute("cy")?.Value;
            if (cx != null && cy != null && long.TryParse(cx, out long width) && long.TryParse(cy, out long height))
            {
                transform.Size = new Size(width, height);
            }
        }

        var rot = xfrm.Attribute("rot")?.Value;
        if (!string.IsNullOrEmpty(rot) && double.TryParse(rot, out double rotation))
        {
            transform.Rotation = rotation / 60000.0; // Convert from 60000ths of a degree
        }

        return transform;
    }

    /// <summary>
    /// Extract text content from txBody element
    /// </summary>
    private string? ExtractTextFromTxBody(XElement? txBody)
    {
        if (txBody == null)
            return null;

        var paragraphs = txBody.Elements(NamespaceConstants.A + "p");
        var textParts = new List<string>();

        foreach (var paragraph in paragraphs)
        {
            var paragraphText = new List<string>();
            var runs = paragraph.Elements(NamespaceConstants.A + "r");

            foreach (var run in runs)
            {
                var textElement = run.Element(NamespaceConstants.A + "t");
                if (textElement != null)
                {
                    paragraphText.Add(textElement.Value);
                }
            }

            if (paragraphText.Count > 0)
            {
                textParts.Add(string.Join("", paragraphText));
            }
        }

        if (textParts.Count == 0)
            return null;

        return string.Join("\n", textParts);
    }

    /// <summary>
    /// Resolve image path from relationship
    /// </summary>
    private string? ResolveImagePath(PackagePart drawingPart, string embedId, Package package)
    {
        try
        {
            var imageRel = drawingPart.GetRelationship(embedId);
            if (imageRel == null)
                return null;

            var imagePart = package.GetPart(
                PackUriHelper.ResolvePartUri(drawingPart.Uri, imageRel.TargetUri));

            // Return the relative path of the image
            return imagePart.Uri.ToString().TrimStart('/');
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to resolve image path for embed ID {EmbedId}", embedId);
            return null;
        }
    }

    private SheetElement? ProcessCell(
        XElement cell,
        Dictionary<int, string> sharedStrings,
        HashSet<int> dateFormats,
        Dictionary<string, string> mergeCells,
        int sheetNumber,
        string sheetName,
        ref int elementIndex)
    {
        var cellRef = cell.Attribute("r")?.Value;
        if (string.IsNullOrEmpty(cellRef))
            return null;

        var column = GetColumnFromCellReference(cellRef);

        var cellType = cell.Attribute("t")?.Value;
        var styleIndex = int.Parse(cell.Attribute("s")?.Value ?? "0");
        var valueElement = cell.Element(NamespaceConstants.spreadsheet + "v");
        var formulaElement = cell.Element(NamespaceConstants.spreadsheet + "f");

        if (valueElement == null && formulaElement == null)
            return null;

        var cellValue = new CellValue();
        var dataType = cellType;

        if (valueElement != null)
        {
            var value = valueElement.Value;

            switch (cellType)
            {
                case "s": // Shared string
                    if (int.TryParse(value, out int stringIndex) && sharedStrings.ContainsKey(stringIndex))
                    {
                        cellValue.Text = sharedStrings[stringIndex];
                        cellValue.ValueType = "text";
                    }
                    break;

                case "b": // Boolean
                    cellValue.Boolean = value == "1";
                    cellValue.ValueType = "boolean";
                    break;

                case "e": // Error
                    cellValue.Text = value;
                    cellValue.ValueType = "error";
                    break;

                case "str": // String (inline)
                case "inlineStr":
                    cellValue.Text = value;
                    cellValue.ValueType = "text";
                    break;

                case "n": // Number (explicit)
                default:
                    if (double.TryParse(value, out double numberValue))
                    {
                        // Check if it's a date based on format
                        if (dateFormats.Contains(styleIndex))
                        {
                            cellValue.Date = ConvertExcelDateToDateTime(numberValue);
                            cellValue.ValueType = "date";
                        }
                        else
                        {
                            cellValue.Number = numberValue;
                            cellValue.ValueType = "number";

                            // Also store text representation for compatibility
                            if (numberValue == Math.Floor(numberValue))
                                cellValue.Text = ((int)numberValue).ToString();
                            else
                                cellValue.Text = numberValue.ToString();
                        }
                    }
                    else
                    {
                        cellValue.Text = value;
                        cellValue.ValueType = "text";
                    }
                    break;
            }
        }

        var element = new SheetElement
        {
            SheetNumber = sheetNumber,
            SheetName = sheetName,
            ElementType = "cell",
            ElementIndex = elementIndex++,
            CellReference = cellRef,
            Value = cellValue.Text != null || cellValue.Number != null || cellValue.Boolean != null || cellValue.Date != null ? cellValue : null,
            Formula = formulaElement?.Value,
            DataType = dataType,
            Format = new CellFormat
            {
                StyleIndex = styleIndex,
                NumFmtId = 0, // Would need styles.xml parsing for accurate value
                IsDate = dateFormats.Contains(styleIndex)
            },
            Column = column
        };

        // Check if this cell is part of a merge range
        if (mergeCells.ContainsKey(cellRef))
        {
            var mergeRange = mergeCells[cellRef];
            var parentCell = GetMergeParentCell(mergeRange);

            element.IsMerged = true;
            element.MergeRange = mergeRange;

            // If this is not the parent cell, set the parent reference
            if (cellRef != parentCell)
            {
                element.MergeParent = parentCell;
            }
        }

        return element;
    }

    private int GetColumnFromCellReference(string cellRef)
    {
        var column = 0;
        foreach (char c in cellRef)
        {
            if (char.IsLetter(c))
            {
                column = column * 26 + (c - 'A' + 1);
            }
            else
            {
                break;
            }
        }
        return column;
    }

    private DateTime ConvertExcelDateToDateTime(double excelDate)
    {
        // Excel dates start from 1900-01-01 (with a leap year bug for 1900)
        if (excelDate < 60)
            return new DateTime(1900, 1, 1).AddDays(excelDate - 1);
        else
            return new DateTime(1900, 1, 1).AddDays(excelDate - 2);
    }

    private string SanitizeFileName(string fileName)
    {
        var invalidChars = Path.GetInvalidFileNameChars()
            .Concat(new[] { '/', '\\', '*', '?', '"', '<', '>', '|', ':' })
            .Distinct()
            .ToArray();

        var sanitized = fileName;
        foreach (var c in invalidChars)
        {
            sanitized = sanitized.Replace(c, '_');
        }

        // Remove leading/trailing spaces and dots
        sanitized = sanitized.Trim(' ', '.');

        // If empty or only underscores, use a default name
        if (string.IsNullOrWhiteSpace(sanitized) || sanitized.All(c => c == '_'))
        {
            sanitized = "Sheet";
        }

        return sanitized;
    }

    /// <summary>
    /// Load merge cells information from worksheet
    /// </summary>
    private Dictionary<string, string> LoadMergeCells(XDocument worksheetDoc)
    {
        var mergeCellsDict = new Dictionary<string, string>();

        var mergeCells = worksheetDoc.Root?.Element(NamespaceConstants.spreadsheet + "mergeCells");
        if (mergeCells == null)
            return mergeCellsDict;

        foreach (var mergeCell in mergeCells.Elements(NamespaceConstants.spreadsheet + "mergeCell"))
        {
            var range = mergeCell.Attribute("ref")?.Value;
            if (string.IsNullOrEmpty(range))
                continue;

            // Parse the range (e.g., "A1:C3")
            var parts = range.Split(':');
            if (parts.Length != 2)
                continue;

            var startCell = parts[0];
            var endCell = parts[1];

            // Get all cells in the range
            var cellsInRange = GetCellsInRange(startCell, endCell);
            foreach (var cell in cellsInRange)
            {
                mergeCellsDict[cell] = range;
            }
        }

        return mergeCellsDict;
    }

    /// <summary>
    /// Get the parent cell (top-left) of a merge range
    /// </summary>
    private string GetMergeParentCell(string mergeRange)
    {
        var parts = mergeRange.Split(':');
        return parts.Length > 0 ? parts[0] : mergeRange;
    }

    /// <summary>
    /// Get all cell references within a range
    /// </summary>
    private List<string> GetCellsInRange(string startCell, string endCell)
    {
        var cells = new List<string>();

        // Parse start cell
        var (startCol, startRow) = ParseCellReference(startCell);
        var (endCol, endRow) = ParseCellReference(endCell);

        for (int row = startRow; row <= endRow; row++)
        {
            for (int col = startCol; col <= endCol; col++)
            {
                cells.Add(GetCellReference(col, row));
            }
        }

        return cells;
    }

    /// <summary>
    /// Parse a cell reference into column and row numbers
    /// </summary>
    private (int col, int row) ParseCellReference(string cellRef)
    {
        int col = 0;
        int rowIndex = 0;

        // Find where the row number starts
        for (int i = 0; i < cellRef.Length; i++)
        {
            if (char.IsDigit(cellRef[i]))
            {
                rowIndex = i;
                break;
            }
            col = col * 26 + (cellRef[i] - 'A' + 1);
        }

        int row = int.Parse(cellRef.Substring(rowIndex));
        return (col, row);
    }

    /// <summary>
    /// Get cell reference from column and row numbers
    /// </summary>
    private string GetCellReference(int col, int row)
    {
        string colStr = "";
        while (col > 0)
        {
            col--;
            colStr = (char)('A' + col % 26) + colStr;
            col /= 26;
        }
        return colStr + row;
    }

    public int ListSheetNames(string inputPath)
    {
        try
        {
            using var package = Package.Open(inputPath, FileMode.Open, FileAccess.Read);

            var workbookPart = PackageUtilities.GetWorkbookPart(package);
            if (workbookPart == null)
            {
                Console.Error.WriteLine("No workbook part found");
                return 1;
            }

            var workbookDoc = PackageUtilities.GetXDocument(workbookPart);
            var sheets = workbookDoc
                .Descendants(NamespaceConstants.spreadsheet + "sheets")
                .Elements(NamespaceConstants.spreadsheet + "sheet")
                .ToList();

            var sheetNumber = 1;
            foreach (var sheet in sheets)
            {
                var sheetName = sheet.Attribute("name")?.Value ?? $"Sheet{sheetNumber}";
                var sanitizedName = SanitizeFileName(sheetName);

                // Output format: sheetNumber|originalName|sanitizedName
                Console.WriteLine($"{sheetNumber}|{sheetName}|{sanitizedName}");
                sheetNumber++;
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing file: {ex.Message}");
            return 1;
        }
    }
}
