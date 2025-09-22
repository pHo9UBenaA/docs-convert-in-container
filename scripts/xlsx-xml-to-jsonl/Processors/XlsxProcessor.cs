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
            var sheetDataByNumber = await ExtractSheetsAsync(inputPath, cancellationToken);

            if (!sheetDataByNumber.Any())
            {
                _logger.LogWarning("No worksheets found in {Path}", inputPath);
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
                var outputPath = Path.Combine(outputDirectory, $"{baseName}_sheet{sheetNumber}.jsonl");
                await WriteSheetToFile(outputPath, sheetElements, cancellationToken);
                outputPaths.Add(outputPath);
            }

            _logger.LogInformation("Successfully processed {Count} sheets", sheetDataByNumber.Count);
            return new ProcessingResult
            {
                Success = true,
                ItemsProcessed = sheetDataByNumber.Count,
                OutputPath = string.Join(";", outputPaths)
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

    private async Task WriteSheetToFile(
        string outputPath,
        List<SheetElement> sheetElements,
        CancellationToken cancellationToken)
    {
        using var writer = new StreamWriter(outputPath);

        foreach (var element in sheetElements)
        {
            var json = JsonSerializer.Serialize(element, ElementJsonSerializerContext.Default.SheetElement);
            // Decode Unicode escapes for Japanese characters
            json = System.Text.RegularExpressions.Regex.Unescape(json);
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
                _logger.LogWarning("No workbook part found in {Path}", path);
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
                var rId = sheet.Attribute(NamespaceConstants.r + "id")?.Value;

                if (string.IsNullOrEmpty(rId))
                    continue;

                var sheetRelationship = workbookPart.GetRelationship(rId);
                var worksheetPart = package.GetPart(
                    PackUriHelper.ResolvePartUri(workbookPart.Uri, sheetRelationship.TargetUri));

                var worksheetDoc = PackageUtilities.GetXDocument(worksheetPart);
                var sheetElements = ProcessWorksheet(worksheetDoc, sharedStrings, dateFormats, sheetNumber, sheetName);

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
                var cellElement = ProcessCell(cell, sharedStrings, dateFormats, sheetNumber, sheetName, ref elementIndex);
                if (cellElement != null)
                {
                    cellElement.Row = rowNum;
                    elements.Add(cellElement);
                }
            }
        }

        return elements;
    }

    private SheetElement? ProcessCell(
        XElement cell,
        Dictionary<int, string> sharedStrings,
        HashSet<int> dateFormats,
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
}