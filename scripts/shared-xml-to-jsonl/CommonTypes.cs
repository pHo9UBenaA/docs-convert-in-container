using System.Text.Json.Serialization;

namespace SharedXmlToJsonl;

/// <summary>
/// Position in EMU (English Metric Units)
/// </summary>
public record Position(long X, long Y);

/// <summary>
/// Size in EMU (English Metric Units)
/// </summary>
public record Size(long Width, long Height);

/// <summary>
/// Transform information for shapes including position, size, and rotation
/// </summary>
public record Transform(
    Position? Position = null,
    Size? Size = null,
    double? Rotation = null,
    CellAnchor? Anchor = null
);

/// <summary>
/// Cell anchor for Excel drawings (two-cell anchor)
/// </summary>
public record CellAnchor(
    string FromCell,
    int FromCol,
    int FromRow,
    string ToCell,
    int ToCol,
    int ToRow
);

/// <summary>
/// Custom geometry path information
/// </summary>
public record CustomGeometry(
    string? PathData = null,
    string? FillRule = null
);

/// <summary>
/// Table cell information
/// </summary>
public record TableCell(
    int Row,
    int Col,
    string? Text = null,
    int? RowSpan = null,
    int? ColSpan = null
);

/// <summary>
/// Line properties for shape borders
/// </summary>
public record LineProperties(
    string? Color = null,
    long? Width = null,
    string? DashStyle = null,
    string? CompoundLineType = null
);

/// <summary>
/// JSON serialization context for common types
/// </summary>
[JsonSerializable(typeof(Position))]
[JsonSerializable(typeof(Size))]
[JsonSerializable(typeof(Transform))]
[JsonSerializable(typeof(CellAnchor))]
[JsonSerializable(typeof(CustomGeometry))]
[JsonSerializable(typeof(TableCell))]
[JsonSerializable(typeof(LineProperties))]
public partial class CommonJsonContext : JsonSerializerContext
{
}