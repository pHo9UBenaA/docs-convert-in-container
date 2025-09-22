using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace SharedXmlToJsonl.Models;

/// <summary>
/// Base class for all element types
/// </summary>
public abstract class ElementBase
{
    public string Type { get; set; } = string.Empty;
}

/// <summary>
/// Shape element
/// </summary>
public class ShapeElement : ElementBase
{
    public ShapeElement() => Type = "shape";
    public string? Id { get; set; }
    public string? Name { get; set; }
    public string? Text { get; set; }
    public Position? Position { get; set; }
    public Size? Size { get; set; }
}

/// <summary>
/// Table element
/// </summary>
public class TableElement : ElementBase
{
    public TableElement() => Type = "table";
    public List<List<string>> Rows { get; set; } = new();
}

/// <summary>
/// Connector element
/// </summary>
public class ConnectorElement : ElementBase
{
    public ConnectorElement() => Type = "connector";
    public string? Id { get; set; }
    public string? Name { get; set; }
}

/// <summary>
/// Picture element
/// </summary>
public class PictureElement : ElementBase
{
    public PictureElement() => Type = "picture";
    public string? Id { get; set; }
    public string? Name { get; set; }
    public string? Description { get; set; }
    public Position? Position { get; set; }
    public Size? Size { get; set; }
}

/// <summary>
/// Chart element
/// </summary>
public class ChartElement : ElementBase
{
    public ChartElement() => Type = "chart";
    public string? Title { get; set; }
    public string? ChartType { get; set; }
}

/// <summary>
/// Spreadsheet table element for Excel worksheets
/// </summary>
public class SpreadsheetTableElement : ElementBase
{
    public SpreadsheetTableElement() => Type = "table";
    public string? SheetName { get; set; }
    public List<List<object?>> Rows { get; set; } = new();
}

/// <summary>
/// JSON serialization context for element types
/// </summary>
[JsonSerializable(typeof(ElementBase))]
[JsonSerializable(typeof(ShapeElement))]
[JsonSerializable(typeof(TableElement))]
[JsonSerializable(typeof(ConnectorElement))]
[JsonSerializable(typeof(PictureElement))]
[JsonSerializable(typeof(ChartElement))]
[JsonSerializable(typeof(SpreadsheetTableElement))]
[JsonSerializable(typeof(List<List<string>>))]
[JsonSerializable(typeof(List<List<object?>>))]
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.SnakeCaseLower,
    WriteIndented = false,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
internal partial class ElementJsonSerializerContext : JsonSerializerContext
{
}