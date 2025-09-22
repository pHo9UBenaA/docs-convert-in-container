using System.Xml.Linq;

namespace SharedXmlToJsonl.Elements
{
    /// <summary>
    /// Base interface for all document elements.
    /// </summary>
    public interface IElement
    {
        /// <summary>
        /// Gets the element type.
        /// </summary>
        string Type { get; }

        /// <summary>
        /// Gets the element ID.
        /// </summary>
        string? Id { get; }

        /// <summary>
        /// Converts the element to a dictionary representation.
        /// </summary>
        /// <returns>A dictionary representing the element.</returns>
        Dictionary<string, object?> ToDictionary();
    }

    /// <summary>
    /// Interface for shape elements.
    /// </summary>
    public interface IShapeElement : IElement
    {
        /// <summary>
        /// Gets the shape properties.
        /// </summary>
        ShapeProperties? Properties { get; }
    }

    /// <summary>
    /// Interface for table elements.
    /// </summary>
    public interface ITableElement : IElement
    {
        /// <summary>
        /// Gets the number of rows.
        /// </summary>
        int RowCount { get; }

        /// <summary>
        /// Gets the number of columns.
        /// </summary>
        int ColumnCount { get; }
    }

    /// <summary>
    /// Interface for connector elements.
    /// </summary>
    public interface IConnectorElement : IElement
    {
        /// <summary>
        /// Gets the start connection ID.
        /// </summary>
        string? StartConnectionId { get; }

        /// <summary>
        /// Gets the end connection ID.
        /// </summary>
        string? EndConnectionId { get; }
    }

    /// <summary>
    /// Represents shape properties.
    /// </summary>
    public class ShapeProperties
    {
        /// <summary>
        /// Gets or sets the X position.
        /// </summary>
        public long? X { get; set; }

        /// <summary>
        /// Gets or sets the Y position.
        /// </summary>
        public long? Y { get; set; }

        /// <summary>
        /// Gets or sets the width.
        /// </summary>
        public long? Width { get; set; }

        /// <summary>
        /// Gets or sets the height.
        /// </summary>
        public long? Height { get; set; }

        /// <summary>
        /// Gets or sets the rotation angle.
        /// </summary>
        public int? Rotation { get; set; }
    }
}