using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Extensions.Logging;
using SharedXmlToJsonl;
using SharedXmlToJsonl.Factories;
using SharedXmlToJsonl.Interfaces;
using SharedXmlToJsonl.Models;
using SharedXmlToJsonl.Configuration;

namespace PptxXmlToJsonl.Processors;

public class PptxProcessor : IPptxProcessor
{
    private readonly IElementFactory _elementFactory;
    private readonly IJsonWriter _jsonWriter;
    private readonly ILogger<PptxProcessor> _logger;

    public PptxProcessor(
        IElementFactory elementFactory,
        IJsonWriter jsonWriter,
        ILogger<PptxProcessor> logger)
    {
        _elementFactory = elementFactory ?? throw new ArgumentNullException(nameof(elementFactory));
        _jsonWriter = jsonWriter ?? throw new ArgumentNullException(nameof(jsonWriter));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    // Implementation of IDocumentProcessor.ProcessAsync
    public async Task<SharedXmlToJsonl.Models.ProcessingResult> ProcessAsync(
        string inputPath,
        string outputDirectory,
        CancellationToken cancellationToken = default)
    {
        // Use default options when called from IDocumentProcessor interface
        var options = new ProcessingOptions();
        return await ProcessAsync(inputPath, outputDirectory, options, cancellationToken);
    }

    // Implementation of IPptxProcessor.ProcessAsync
    public async Task<SharedXmlToJsonl.Models.ProcessingResult> ProcessAsync(
        string inputPath,
        string outputDirectory,
        ProcessingOptions options,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Starting PPTX processing for {Path}", inputPath);

        try
        {
            var slideDataByNumber = await ExtractSlidesAsync(inputPath, cancellationToken);

            if (!slideDataByNumber.Any())
            {
                _logger.LogWarning("No slides found in {Path}", inputPath);
                return new ProcessingResult
                {
                    Success = false,
                    ErrorMessage = "No slides found in the presentation"
                };
            }

            var baseName = Path.GetFileNameWithoutExtension(inputPath);
            var outputPaths = new List<string>();

            // Write each slide to a separate file
            foreach (var (slideNumber, slideElements) in slideDataByNumber)
            {
                var outputPath = Path.Combine(outputDirectory, $"{baseName}_page-{slideNumber}.jsonl");
                await WriteSlideToFile(outputPath, slideElements, cancellationToken);
                outputPaths.Add(outputPath);
            }

            _logger.LogInformation("Successfully processed {Count} slides", slideDataByNumber.Count);
            return new SharedXmlToJsonl.Models.ProcessingResult
            {
                Success = true,
                ItemsProcessed = slideDataByNumber.Count,
                OutputPath = string.Join(";", outputPaths)
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing PPTX file {Path}", inputPath);
            return new SharedXmlToJsonl.Models.ProcessingResult
            {
                Success = false,
                ErrorMessage = ex.Message
            };
        }
    }

    private async Task WriteSlideToFile(
        string outputPath,
        List<SlideElement> slideElements,
        CancellationToken cancellationToken)
    {
        using var writer = new StreamWriter(outputPath);

        foreach (var element in slideElements)
        {
            var json = JsonSerializer.Serialize(element, ElementJsonSerializerContext.Default.SlideElement);
            await writer.WriteLineAsync(json);
        }
    }

    private async Task<Dictionary<int, List<SlideElement>>> ExtractSlidesAsync(
        string path,
        CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            using var package = Package.Open(path, FileMode.Open, FileAccess.Read);
            var slideDataByNumber = new Dictionary<int, List<SlideElement>>();

            var presentationPart = PackageUtilities.GetPresentationPart(package);
            if (presentationPart == null)
            {
                _logger.LogWarning("No presentation part found in {Path}", path);
                return slideDataByNumber;
            }

            var presentationDoc = PackageUtilities.GetXDocument(presentationPart);
            var slideIds = presentationDoc
                .Descendants(NamespaceConstants.p + "sldIdLst")
                .Elements(NamespaceConstants.p + "sldId")
                .ToList();

            var slideNumber = 1;
            foreach (var slideId in slideIds)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var rId = slideId.Attribute(NamespaceConstants.r + "id")?.Value;
                if (string.IsNullOrEmpty(rId))
                    continue;

                var slideRelationship = presentationPart.GetRelationship(rId);
                var slidePart = package.GetPart(
                    PackUriHelper.ResolvePartUri(presentationPart.Uri, slideRelationship.TargetUri));

                var slideDoc = PackageUtilities.GetXDocument(slidePart);
                var slideElements = ExtractSlideElements(slideDoc, slideNumber);

                slideDataByNumber[slideNumber] = slideElements;
                slideNumber++;
            }

            return slideDataByNumber;
        }, cancellationToken);
    }

    private List<SlideElement> ExtractSlideElements(XDocument slideDoc, int slideNumber)
    {
        var elements = new List<SlideElement>();
        var elementIndex = 0;

        // Add slide metadata as first element
        elements.Add(new SlideElement
        {
            SlideNumber = slideNumber,
            ElementType = "slide_metadata",
            ElementIndex = elementIndex++,
            Metadata = new Dictionary<string, object> { ["slide_number"] = slideNumber }
        });

        var cSld = slideDoc.Root?.Element(NamespaceConstants.p + "cSld");
        if (cSld == null)
            return elements;

        var spTree = cSld.Element(NamespaceConstants.p + "spTree");
        if (spTree == null)
            return elements;

        foreach (var element in spTree.Elements())
        {
            var localName = element.Name.LocalName;

            switch (localName)
            {
                case "sp":  // Shape
                    var shapeElements = ProcessShape(element, slideNumber, ref elementIndex);
                    elements.AddRange(shapeElements);
                    break;

                case "graphicFrame":  // Table, Chart, SmartArt
                    var graphicElements = ProcessGraphicFrame(element, slideNumber, ref elementIndex);
                    elements.AddRange(graphicElements);
                    break;

                case "cxnSp":  // Connector
                    var connectorElement = ProcessConnector(element, slideNumber, ref elementIndex);
                    if (connectorElement != null)
                        elements.Add(connectorElement);
                    break;

                case "pic":  // Picture
                    var pictureElement = ProcessPicture(element, slideNumber, ref elementIndex);
                    if (pictureElement != null)
                        elements.Add(pictureElement);
                    break;
            }
        }

        return elements;
    }

    private List<SlideElement> ProcessShape(XElement shapeElement, int slideNumber, ref int elementIndex)
    {
        var elements = new List<SlideElement>();

        var nvSpPr = shapeElement.Element(NamespaceConstants.p + "nvSpPr");
        var cNvPr = nvSpPr?.Element(NamespaceConstants.p + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";

        var spPr = shapeElement.Element(NamespaceConstants.p + "spPr");
        var transform = ExtractTransform(spPr);

        // Extract shape type
        var prstGeom = spPr?.Element(NamespaceConstants.a + "prstGeom");
        var shapeType = prstGeom?.Attribute("prst")?.Value;

        // Extract line properties
        var lineProperties = ShapeProcessor.ExtractLineProperties(spPr, NamespaceConstants.a);

        // Extract fill information
        var (hasFill, fillColor) = ShapeProcessor.ExtractFillInfo(spPr, NamespaceConstants.a);

        // Add shape element
        var shapeEl = new SlideElement
        {
            SlideNumber = slideNumber,
            ElementType = "shape",
            ElementIndex = elementIndex++,
            ShapeId = id,
            ShapeName = name,
            Transform = transform,
            ShapeType = shapeType,
            GroupLevel = 1,
            LineProperties = lineProperties,
            HasFill = hasFill,
            FillColor = fillColor
        };
        elements.Add(shapeEl);

        // Extract text content - each paragraph as separate element
        var txBody = shapeElement.Element(NamespaceConstants.p + "txBody");
        if (txBody != null)
        {
            foreach (var paragraph in txBody.Elements(NamespaceConstants.a + "p"))
            {
                var paragraphText = ExtractParagraphText(paragraph);
                if (!string.IsNullOrEmpty(paragraphText))
                {
                    var textEl = new SlideElement
                    {
                        SlideNumber = slideNumber,
                        ElementType = "text",
                        ElementIndex = elementIndex++,
                        Text = paragraphText,
                        ShapeId = id,
                        ShapeName = name,
                        Transform = transform,
                        GroupLevel = 1
                    };
                    elements.Add(textEl);
                }
            }
        }

        return elements;
    }

    private SharedXmlToJsonl.Models.Transform? ExtractTransform(XElement? spPr)
    {
        if (spPr == null)
            return null;

        var xfrm = spPr.Element(NamespaceConstants.a + "xfrm");
        if (xfrm == null)
            return null;

        var transform = new SharedXmlToJsonl.Models.Transform();

        var off = xfrm.Element(NamespaceConstants.a + "off");
        if (off != null)
        {
            var x = off.Attribute("x")?.Value;
            var y = off.Attribute("y")?.Value;
            if (x != null && y != null && int.TryParse(x, out int xVal) && int.TryParse(y, out int yVal))
            {
                transform.Position = new Position(xVal, yVal);
            }
        }

        var ext = xfrm.Element(NamespaceConstants.a + "ext");
        if (ext != null)
        {
            var cx = ext.Attribute("cx")?.Value;
            var cy = ext.Attribute("cy")?.Value;
            if (cx != null && cy != null && int.TryParse(cx, out int width) && int.TryParse(cy, out int height))
            {
                transform.Size = new Size(width, height);
            }
        }

        return transform;
    }

    private List<SlideElement> ProcessGraphicFrame(XElement graphicFrame, int slideNumber, ref int elementIndex)
    {
        var elements = new List<SlideElement>();

        var nvGraphicFramePr = graphicFrame.Element(NamespaceConstants.p + "nvGraphicFramePr");
        var cNvPr = nvGraphicFramePr?.Element(NamespaceConstants.p + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";

        var xfrm = graphicFrame.Element(NamespaceConstants.p + "xfrm");
        var transform = ExtractTransformFromXfrm(xfrm);

        var graphic = graphicFrame.Element(NamespaceConstants.a + "graphic");
        var graphicData = graphic?.Element(NamespaceConstants.a + "graphicData");

        if (graphicData == null)
            return elements;

        var tbl = graphicData.Element(NamespaceConstants.a + "tbl");
        if (tbl != null)
        {
            var tableElement = new SlideElement
            {
                SlideNumber = slideNumber,
                ElementType = "table",
                ElementIndex = elementIndex++,
                ShapeId = id,
                ShapeName = name,
                Transform = transform,
                ShapeType = "table",
                GroupLevel = 1
            };
            elements.Add(tableElement);

            // Process table rows
            foreach (var tr in tbl.Elements(NamespaceConstants.a + "tr"))
            {
                foreach (var tc in tr.Elements(NamespaceConstants.a + "tc"))
                {
                    var txBody = tc.Element(NamespaceConstants.a + "txBody");
                    var cellText = ExtractTextContent(txBody);
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        var cellElement = new SlideElement
                        {
                            SlideNumber = slideNumber,
                            ElementType = "table_cell",
                            ElementIndex = elementIndex++,
                            Text = cellText,
                            ShapeId = id,
                            ShapeName = name,
                            Transform = transform,
                            GroupLevel = 2
                        };
                        elements.Add(cellElement);
                    }
                }
            }
        }

        return elements;
    }

    private SharedXmlToJsonl.Models.Transform? ExtractTransformFromXfrm(XElement? xfrm)
    {
        if (xfrm == null)
            return null;

        var transform = new SharedXmlToJsonl.Models.Transform();

        var off = xfrm.Element(NamespaceConstants.a + "off");
        if (off != null)
        {
            var x = off.Attribute("x")?.Value;
            var y = off.Attribute("y")?.Value;
            if (x != null && y != null && int.TryParse(x, out int xVal) && int.TryParse(y, out int yVal))
            {
                transform.Position = new Position(xVal, yVal);
            }
        }

        var ext = xfrm.Element(NamespaceConstants.a + "ext");
        if (ext != null)
        {
            var cx = ext.Attribute("cx")?.Value;
            var cy = ext.Attribute("cy")?.Value;
            if (cx != null && cy != null && int.TryParse(cx, out int width) && int.TryParse(cy, out int height))
            {
                transform.Size = new Size(width, height);
            }
        }

        return transform;
    }

    private SlideElement? ProcessConnector(XElement connectorElement, int slideNumber, ref int elementIndex)
    {
        var nvCxnSpPr = connectorElement.Element(NamespaceConstants.p + "nvCxnSpPr");
        var cNvPr = nvCxnSpPr?.Element(NamespaceConstants.p + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";

        var spPr = connectorElement.Element(NamespaceConstants.p + "spPr");
        var transform = ExtractTransform(spPr);

        return new SlideElement
        {
            SlideNumber = slideNumber,
            ElementType = "connector",
            ElementIndex = elementIndex++,
            ShapeId = id,
            ShapeName = name,
            Transform = transform,
            ShapeType = "connector",
            GroupLevel = 1
        };
    }

    private SlideElement? ProcessPicture(XElement pictureElement, int slideNumber, ref int elementIndex)
    {
        var nvPicPr = pictureElement.Element(NamespaceConstants.p + "nvPicPr");
        var cNvPr = nvPicPr?.Element(NamespaceConstants.p + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";
        var descr = cNvPr?.Attribute("descr")?.Value;

        var spPr = pictureElement.Element(NamespaceConstants.p + "spPr");
        var transform = ExtractTransform(spPr);

        return new SlideElement
        {
            SlideNumber = slideNumber,
            ElementType = "picture",
            ElementIndex = elementIndex++,
            ShapeId = id,
            ShapeName = name,
            Text = descr,
            Transform = transform,
            ShapeType = "picture",
            GroupLevel = 1
        };
    }

    private string ExtractParagraphText(XElement paragraph)
    {
        var paragraphText = new List<string>();
        foreach (var run in paragraph.Elements(NamespaceConstants.a + "r"))
        {
            var text = run.Element(NamespaceConstants.a + "t")?.Value;
            if (!string.IsNullOrEmpty(text))
                paragraphText.Add(text);
        }

        return string.Join("", paragraphText);
    }

    private string ExtractTextContent(XElement? txBody)
    {
        if (txBody == null)
            return string.Empty;

        var texts = new List<string>();
        foreach (var paragraph in txBody.Elements(NamespaceConstants.a + "p"))
        {
            var paragraphText = ExtractParagraphText(paragraph);
            if (!string.IsNullOrEmpty(paragraphText))
                texts.Add(paragraphText);
        }

        return string.Join("\n", texts);
    }
}