using System.Collections.Generic;
using System.Text.Json.Serialization;
using SharedXmlToJsonl;

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
/// Slide element for PPTX presentations
/// </summary>
public class SlideElement
{
    [JsonPropertyName("slide_number")]
    public int SlideNumber { get; set; }

    [JsonPropertyName("element_type")]
    public string ElementType { get; set; } = string.Empty;

    [JsonPropertyName("element_index")]
    public int ElementIndex { get; set; }

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("shape_id")]
    public string? ShapeId { get; set; }

    [JsonPropertyName("shape_name")]
    public string? ShapeName { get; set; }

    [JsonPropertyName("transform")]
    public Transform? Transform { get; set; }

    [JsonPropertyName("shape_type")]
    public string? ShapeType { get; set; }

    [JsonPropertyName("group_level")]
    public int? GroupLevel { get; set; }

    [JsonPropertyName("parent_group_id")]
    public string? ParentGroupId { get; set; }

    [JsonPropertyName("custom_geometry")]
    public object? CustomGeometry { get; set; }

    [JsonPropertyName("ole_object_type")]
    public string? OleObjectType { get; set; }

    [JsonPropertyName("content_part_ref")]
    public string? ContentPartRef { get; set; }

    [JsonPropertyName("line_properties")]
    public object? LineProperties { get; set; }

    [JsonPropertyName("has_fill")]
    public bool? HasFill { get; set; }

    [JsonPropertyName("fill_color")]
    public string? FillColor { get; set; }

    [JsonPropertyName("metadata")]
    public object? Metadata { get; set; }

    [JsonPropertyName("error_info")]
    public string? ErrorInfo { get; set; }
}

/// <summary>
/// Transform information for shapes
/// </summary>
public class Transform
{
    [JsonPropertyName("position")]
    public Position? Position { get; set; }

    [JsonPropertyName("size")]
    public Size? Size { get; set; }

    [JsonPropertyName("rotation")]
    public double? Rotation { get; set; }

    [JsonPropertyName("anchor")]
    public object? Anchor { get; set; }
}

/// <summary>
/// Sheet element for XLSX spreadsheets
/// </summary>
public class SheetElement
{
    [JsonPropertyName("sheet_number")]
    public int SheetNumber { get; set; }

    [JsonPropertyName("sheet_name")]
    public string SheetName { get; set; } = string.Empty;

    [JsonPropertyName("element_type")]
    public string ElementType { get; set; } = string.Empty;

    [JsonPropertyName("element_index")]
    public int ElementIndex { get; set; }

    [JsonPropertyName("cell_reference")]
    public string? CellReference { get; set; }

    [JsonPropertyName("value")]
    public CellValue? Value { get; set; }

    [JsonPropertyName("formula")]
    public string? Formula { get; set; }

    [JsonPropertyName("data_type")]
    public string? DataType { get; set; }

    [JsonPropertyName("format")]
    public CellFormat? Format { get; set; }

    [JsonPropertyName("row")]
    public int? Row { get; set; }

    [JsonPropertyName("column")]
    public int? Column { get; set; }

    [JsonPropertyName("transform")]
    public Transform? Transform { get; set; }

    [JsonPropertyName("shape_type")]
    public string? ShapeType { get; set; }

    [JsonPropertyName("shape_id")]
    public string? ShapeId { get; set; }

    [JsonPropertyName("shape_name")]
    public string? ShapeName { get; set; }

    [JsonPropertyName("group_level")]
    public int? GroupLevel { get; set; }

    [JsonPropertyName("parent_group_id")]
    public string? ParentGroupId { get; set; }

    [JsonPropertyName("custom_geometry")]
    public object? CustomGeometry { get; set; }

    [JsonPropertyName("ole_object_type")]
    public string? OleObjectType { get; set; }

    [JsonPropertyName("metadata")]
    public object? Metadata { get; set; }

    [JsonPropertyName("error_info")]
    public string? ErrorInfo { get; set; }

    [JsonPropertyName("merge_range")]
    public string? MergeRange { get; set; }

    [JsonPropertyName("is_merged")]
    public bool? IsMerged { get; set; }

    [JsonPropertyName("merge_parent")]
    public string? MergeParent { get; set; }

    [JsonPropertyName("image_path")]
    public string? ImagePath { get; set; }

    [JsonPropertyName("anchor_type")]
    public string? AnchorType { get; set; }

    [JsonPropertyName("anchor_from")]
    public CellAnchor? AnchorFrom { get; set; }

    [JsonPropertyName("anchor_to")]
    public CellAnchor? AnchorTo { get; set; }
}

/// <summary>
/// Cell value for spreadsheet cells
/// </summary>
public class CellValue
{
    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("number")]
    public double? Number { get; set; }

    [JsonPropertyName("boolean")]
    public bool? Boolean { get; set; }

    [JsonPropertyName("date")]
    public DateTime? Date { get; set; }

    [JsonPropertyName("value_type")]
    public string ValueType { get; set; } = string.Empty;
}

/// <summary>
/// Cell format information
/// </summary>
public class CellFormat
{
    [JsonPropertyName("style_index")]
    public int StyleIndex { get; set; }

    [JsonPropertyName("num_fmt_id")]
    public int NumFmtId { get; set; }

    [JsonPropertyName("num_fmt_code")]
    public string? NumFmtCode { get; set; }

    [JsonPropertyName("is_date")]
    public bool IsDate { get; set; }
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
[JsonSerializable(typeof(SlideElement))]
[JsonSerializable(typeof(SheetElement))]
[JsonSerializable(typeof(Transform))]
[JsonSerializable(typeof(Position))]
[JsonSerializable(typeof(Size))]
[JsonSerializable(typeof(CellValue))]
[JsonSerializable(typeof(CellFormat))]
[JsonSerializable(typeof(LineProperties))]
[JsonSerializable(typeof(CellAnchor))]
[JsonSerializable(typeof(List<List<string>>))]
[JsonSerializable(typeof(List<List<object?>>))]
[JsonSerializable(typeof(Dictionary<string, object>))]
[JsonSerializable(typeof(Dictionary<string, int>))]
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.SnakeCaseLower,
    WriteIndented = false)]
public partial class ElementJsonSerializerContext : JsonSerializerContext
{
}