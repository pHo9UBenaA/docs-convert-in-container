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

namespace XlsxXmlToJsonl.Processors;

public class XlsxProcessor : IXlsxProcessor
{
    private readonly IElementFactory _elementFactory;
    private readonly IJsonWriter _jsonWriter;
    private readonly ILogger<XlsxProcessor> _logger;

    public XlsxProcessor(
        IElementFactory elementFactory,
        IJsonWriter jsonWriter,
        ILogger<XlsxProcessor> logger)
    {
        _elementFactory = elementFactory ?? throw new ArgumentNullException(nameof(elementFactory));
        _jsonWriter = jsonWriter ?? throw new ArgumentNullException(nameof(jsonWriter));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
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
        _logger.LogInformation("Starting XLSX processing for {Path}", inputPath);

        try
        {
            var entries = await ExtractEntriesAsync(inputPath, cancellationToken);

            if (!entries.Any())
            {
                _logger.LogWarning("No worksheets found in {Path}", inputPath);
                return new ProcessingResult
                {
                    Success = false,
                    ErrorMessage = "No worksheets found in the workbook"
                };
            }

            var outputPath = Path.Combine(outputDirectory,
                Path.GetFileNameWithoutExtension(inputPath) + ".jsonl");

            await _jsonWriter.WriteJsonLinesAsync(outputPath, entries, cancellationToken);

            _logger.LogInformation("Successfully processed {Count} entries", entries.Count);
            return new ProcessingResult
            {
                Success = true,
                ProcessedItems = entries.Count,
                OutputPath = outputPath
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing XLSX file {Path}", inputPath);
            return new ProcessingResult
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

            // Load shared strings
            var sharedStrings = LoadSharedStrings(package);

            // Get workbook part
            var workbookPart = PackageUtilities.GetWorkbookPart(package);
            if (workbookPart == null)
            {
                _logger.LogWarning("No workbook part found in {Path}", path);
                return entries;
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
                var rId = sheet.Attribute(NamespaceConstants.r + "id")?.Value;

                if (string.IsNullOrEmpty(rId))
                    continue;

                var sheetRelationship = workbookPart.GetRelationship(rId);
                var worksheetPart = package.GetPart(
                    PackUriHelper.ResolvePartUri(workbookPart.Uri, sheetRelationship.TargetUri));

                var worksheetDoc = PackageUtilities.GetXDocument(worksheetPart);
                var worksheetData = ProcessWorksheet(worksheetDoc, sharedStrings, sheetName);

                foreach (var element in worksheetData.Elements)
                {
                    entries.Add(new DocumentEntry
                    {
                        Document = DocumentUtilities.GetDocumentNameFromPath(path),
                        DocumentType = "xlsx",
                        PageNumber = sheetNumber,
                        SheetName = sheetName,
                        Element = element
                    });
                }

                sheetNumber++;
            }

            return entries;
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

    private WorksheetData ProcessWorksheet(
        XDocument worksheetDoc,
        Dictionary<int, string> sharedStrings,
        string sheetName)
    {
        var data = new WorksheetData { SheetName = sheetName };
        var sheetData = worksheetDoc.Root?.Element(NamespaceConstants.spreadsheet + "sheetData");

        if (sheetData == null)
            return data;

        var rows = new List<List<object?>>();
        foreach (var row in sheetData.Elements(NamespaceConstants.spreadsheet + "row"))
        {
            var rowData = ProcessRow(row, sharedStrings);
            if (rowData.Any(cell => cell != null))
            {
                rows.Add(rowData);
            }
        }

        if (rows.Any())
        {
            data.Elements.Add(new
            {
                Type = "table",
                SheetName = sheetName,
                Rows = rows
            });
        }

        // Process shapes/drawings
        var drawing = worksheetDoc.Root?.Element(NamespaceConstants.spreadsheet + "drawing");
        if (drawing != null)
        {
            var shapes = ProcessDrawing(drawing);
            data.Elements.AddRange(shapes);
        }

        return data;
    }

    private List<object?> ProcessRow(XElement row, Dictionary<int, string> sharedStrings)
    {
        var cells = new List<object?>();
        var rowNum = row.Attribute("r")?.Value;

        foreach (var cell in row.Elements(NamespaceConstants.spreadsheet + "c"))
        {
            var cellValue = GetCellValue(cell, sharedStrings);
            cells.Add(cellValue);
        }

        return cells;
    }

    private object? GetCellValue(XElement cell, Dictionary<int, string> sharedStrings)
    {
        var cellType = cell.Attribute("t")?.Value;
        var valueElement = cell.Element(NamespaceConstants.spreadsheet + "v");
        var formulaElement = cell.Element(NamespaceConstants.spreadsheet + "f");

        if (valueElement == null)
            return null;

        var value = valueElement.Value;

        // Handle different cell types
        switch (cellType)
        {
            case "s": // Shared string
                if (int.TryParse(value, out int stringIndex) && sharedStrings.ContainsKey(stringIndex))
                {
                    return sharedStrings[stringIndex];
                }
                break;

            case "b": // Boolean
                return value == "1";

            case "e": // Error
                return $"#ERROR: {value}";

            case "str": // String (inline)
                return value;

            case "n": // Number (default)
            default:
                if (double.TryParse(value, out double numberValue))
                {
                    // Check if it's likely a date (Excel dates are numbers)
                    if (IsLikelyDate(cell))
                    {
                        return ConvertExcelDateToDateTime(numberValue);
                    }
                    return numberValue;
                }
                return value;
        }

        return value;
    }

    private bool IsLikelyDate(XElement cell)
    {
        // Check if cell has a style that indicates it's a date
        var styleIndex = cell.Attribute("s")?.Value;
        // This is a simplified check - in a full implementation,
        // you would check the style definitions in styles.xml
        return false;
    }

    private DateTime ConvertExcelDateToDateTime(double excelDate)
    {
        // Excel dates start from 1900-01-01 (with a leap year bug for 1900)
        if (excelDate < 60)
            return new DateTime(1900, 1, 1).AddDays(excelDate - 1);
        else
            return new DateTime(1900, 1, 1).AddDays(excelDate - 2);
    }

    private List<object> ProcessDrawing(XElement drawing)
    {
        var shapes = new List<object>();
        // Simplified - would need to follow relationship to drawing part
        // and process the shapes there
        return shapes;
    }

    private class WorksheetData
    {
        public string SheetName { get; set; } = "";
        public List<object> Elements { get; set; } = new();
    }
}