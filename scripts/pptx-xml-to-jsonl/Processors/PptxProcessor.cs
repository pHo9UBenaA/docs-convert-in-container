using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
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
            var entries = await ExtractEntriesAsync(inputPath, cancellationToken);

            if (!entries.Any())
            {
                _logger.LogWarning("No slides found in {Path}", inputPath);
                return new ProcessingResult
                {
                    Success = false,
                    ErrorMessage = "No slides found in the presentation"
                };
            }

            var outputPath = Path.Combine(outputDirectory,
                Path.GetFileNameWithoutExtension(inputPath) + ".jsonl");

            await _jsonWriter.WriteJsonLinesAsync(outputPath, entries, cancellationToken);

            _logger.LogInformation("Successfully processed {Count} slides", entries.Count);
            return new SharedXmlToJsonl.Models.ProcessingResult
            {
                Success = true,
                ItemsProcessed = entries.Count,
                OutputPath = outputPath
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

    private async Task<IReadOnlyList<DocumentEntry>> ExtractEntriesAsync(
        string path,
        CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            using var package = Package.Open(path, FileMode.Open, FileAccess.Read);
            var entries = new List<DocumentEntry>();
            var slideNumber = 1;

            var presentationPart = PackageUtilities.GetPresentationPart(package);
            if (presentationPart == null)
            {
                _logger.LogWarning("No presentation part found in {Path}", path);
                return entries;
            }

            var presentationDoc = PackageUtilities.GetXDocument(presentationPart);
            var slideIds = presentationDoc
                .Descendants(NamespaceConstants.p + "sldIdLst")
                .Elements(NamespaceConstants.p + "sldId")
                .ToList();

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
                var elements = ExtractSlideElements(slideDoc, slideNumber);

                foreach (var element in elements)
                {
                    entries.Add(new DocumentEntry
                    {
                        Document = DocumentUtilities.GetDocumentNameFromPath(path),
                        DocumentType = "pptx",
                        PageNumber = slideNumber,
                        Element = element
                    });
                }

                slideNumber++;
            }

            return entries;
        }, cancellationToken);
    }

    private IReadOnlyList<dynamic> ExtractSlideElements(XDocument slideDoc, int slideNumber)
    {
        var elements = new List<dynamic>();
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
                    var shape = ProcessShape(element);
                    if (shape != null)
                        elements.Add(shape);
                    break;

                case "graphicFrame":  // Table, Chart, SmartArt
                    var graphic = ProcessGraphicFrame(element);
                    if (graphic != null)
                        elements.Add(graphic);
                    break;

                case "cxnSp":  // Connector
                    var connector = ProcessConnector(element);
                    if (connector != null)
                        elements.Add(connector);
                    break;

                case "pic":  // Picture
                    var picture = ProcessPicture(element);
                    if (picture != null)
                        elements.Add(picture);
                    break;
            }
        }

        return elements;
    }

    private object? ProcessShape(XElement shapeElement)
    {
        var nvSpPr = shapeElement.Element(NamespaceConstants.p + "nvSpPr");
        var cNvPr = nvSpPr?.Element(NamespaceConstants.p + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";

        var spPr = shapeElement.Element(NamespaceConstants.p + "spPr");
        var position = XmlUtilities.GetPositionFromTransform(spPr, NamespaceConstants.a);
        var size = XmlUtilities.GetSizeFromTransform(spPr, NamespaceConstants.a);

        var txBody = shapeElement.Element(NamespaceConstants.p + "txBody");
        var text = ExtractTextContent(txBody);

        if (string.IsNullOrEmpty(text) && position == null && size == null)
            return null;

        return new
        {
            Type = "shape",
            Id = id,
            Name = name,
            Text = text,
            Position = position,
            Size = size
        };
    }

    private object? ProcessGraphicFrame(XElement graphicFrame)
    {
        var graphic = graphicFrame.Element(NamespaceConstants.a + "graphic");
        var graphicData = graphic?.Element(NamespaceConstants.a + "graphicData");

        if (graphicData == null)
            return null;

        var tbl = graphicData.Element(NamespaceConstants.a + "tbl");
        if (tbl != null)
        {
            return ProcessTable(tbl);
        }

        return null;
    }

    private object? ProcessTable(XElement tableElement)
    {
        var rows = new List<List<string>>();

        foreach (var tr in tableElement.Elements(NamespaceConstants.a + "tr"))
        {
            var row = new List<string>();
            foreach (var tc in tr.Elements(NamespaceConstants.a + "tc"))
            {
                var txBody = tc.Element(NamespaceConstants.a + "txBody");
                var cellText = ExtractTextContent(txBody);
                row.Add(cellText);
            }
            rows.Add(row);
        }

        if (!rows.Any())
            return null;

        return new
        {
            Type = "table",
            Rows = rows
        };
    }

    private object? ProcessConnector(XElement connectorElement)
    {
        var nvCxnSpPr = connectorElement.Element(NamespaceConstants.p + "nvCxnSpPr");
        var cNvPr = nvCxnSpPr?.Element(NamespaceConstants.p + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";

        return new
        {
            Type = "connector",
            Id = id,
            Name = name
        };
    }

    private object? ProcessPicture(XElement pictureElement)
    {
        var nvPicPr = pictureElement.Element(NamespaceConstants.p + "nvPicPr");
        var cNvPr = nvPicPr?.Element(NamespaceConstants.p + "cNvPr");
        var id = cNvPr?.Attribute("id")?.Value ?? "";
        var name = cNvPr?.Attribute("name")?.Value ?? "";
        var descr = cNvPr?.Attribute("descr")?.Value;

        var spPr = pictureElement.Element(NamespaceConstants.p + "spPr");
        var position = XmlUtilities.GetPositionFromTransform(spPr, NamespaceConstants.a);
        var size = XmlUtilities.GetSizeFromTransform(spPr, NamespaceConstants.a);

        return new
        {
            Type = "picture",
            Id = id,
            Name = name,
            Description = descr,
            Position = position,
            Size = size
        };
    }

    private string ExtractTextContent(XElement? txBody)
    {
        if (txBody == null)
            return string.Empty;

        var texts = new List<string>();
        foreach (var paragraph in txBody.Elements(NamespaceConstants.a + "p"))
        {
            var paragraphText = new List<string>();
            foreach (var run in paragraph.Elements(NamespaceConstants.a + "r"))
            {
                var text = run.Element(NamespaceConstants.a + "t")?.Value;
                if (!string.IsNullOrEmpty(text))
                    paragraphText.Add(text);
            }

            if (paragraphText.Any())
                texts.Add(string.Join("", paragraphText));
        }

        return string.Join("\n", texts);
    }
}