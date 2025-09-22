using System.Xml.Linq;
using SharedXmlToJsonl.Elements;

namespace SharedXmlToJsonl.Factories
{
    /// <summary>
    /// Factory interface for creating document elements.
    /// </summary>
    public interface IElementFactory
    {
        /// <summary>
        /// Creates an element from an XML element.
        /// </summary>
        /// <param name="elementType">The type of element to create.</param>
        /// <param name="xmlElement">The XML element to parse.</param>
        /// <returns>The created element.</returns>
        IElement CreateElement(string elementType, XElement xmlElement);

        /// <summary>
        /// Creates a shape element from an XML element.
        /// </summary>
        /// <param name="shapeElement">The XML shape element.</param>
        /// <returns>The created shape element.</returns>
        IShapeElement CreateShape(XElement shapeElement);

        /// <summary>
        /// Creates a table element from an XML element.
        /// </summary>
        /// <param name="tableElement">The XML table element.</param>
        /// <returns>The created table element.</returns>
        ITableElement CreateTable(XElement tableElement);

        /// <summary>
        /// Creates a connector element from an XML element.
        /// </summary>
        /// <param name="connectorElement">The XML connector element.</param>
        /// <returns>The created connector element.</returns>
        IConnectorElement CreateConnector(XElement connectorElement);
    }
}