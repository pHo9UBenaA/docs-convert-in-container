using System;
using System.Collections.Generic;
using System.Xml.Linq;
using SharedXmlToJsonl.Elements;

namespace SharedXmlToJsonl.Factories
{
    /// <summary>
    /// Factory implementation for creating document elements from XML.
    /// </summary>
    public class ElementFactory : IElementFactory
    {
        private readonly XNamespace _presentationNamespace;
        private readonly XNamespace _drawingNamespace;

        /// <summary>
        /// Initializes a new instance of the ElementFactory class.
        /// </summary>
        public ElementFactory()
        {
            _presentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
            _drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        }

        /// <summary>
        /// Creates an element based on type.
        /// </summary>
        public IElement CreateElement(string elementType, XElement xmlElement)
        {
            return elementType.ToLowerInvariant() switch
            {
                "shape" => CreateShape(xmlElement),
                "table" => CreateTable(xmlElement),
                "connector" => CreateConnector(xmlElement),
                _ => new GenericElement(elementType, xmlElement)
            };
        }

        /// <summary>
        /// Creates a shape element from XML.
        /// </summary>
        public IShapeElement CreateShape(XElement shapeElement)
        {
            return new ShapeElement(shapeElement, _presentationNamespace, _drawingNamespace);
        }

        /// <summary>
        /// Creates a table element from XML.
        /// </summary>
        public ITableElement CreateTable(XElement tableElement)
        {
            return new TableElement(tableElement, _drawingNamespace);
        }

        /// <summary>
        /// Creates a connector element from XML.
        /// </summary>
        public IConnectorElement CreateConnector(XElement connectorElement)
        {
            return new ConnectorElement(connectorElement, _presentationNamespace);
        }

        /// <summary>
        /// Generic element implementation.
        /// </summary>
        private sealed class GenericElement : IElement
        {
            public string Type { get; }
            public string? Id { get; }
            private readonly XElement _xmlElement;

            public GenericElement(string type, XElement xmlElement)
            {
                Type = type;
                _xmlElement = xmlElement;
                Id = xmlElement.Attribute("id")?.Value;
            }

            public Dictionary<string, object?> ToDictionary()
            {
                return new Dictionary<string, object?>
                {
                    ["type"] = Type,
                    ["id"] = Id
                };
            }
        }

        /// <summary>
        /// Shape element implementation.
        /// </summary>
        private sealed class ShapeElement : IShapeElement
        {
            public string Type => "shape";
            public string? Id { get; }
            public ShapeProperties? Properties { get; }

            public ShapeElement(XElement xmlElement, XNamespace p, XNamespace a)
            {
                var nvSpPr = xmlElement.Element(p + "nvSpPr");
                var cNvPr = nvSpPr?.Element(p + "cNvPr");
                Id = cNvPr?.Attribute("id")?.Value;

                var spPr = xmlElement.Element(p + "spPr");
                var xfrm = spPr?.Element(a + "xfrm");

                if (xfrm != null)
                {
                    Properties = new ShapeProperties();
                    var off = xfrm.Element(a + "off");
                    var ext = xfrm.Element(a + "ext");

                    if (off != null)
                    {
                        if (long.TryParse(off.Attribute("x")?.Value, out var x)) Properties.X = x;
                        if (long.TryParse(off.Attribute("y")?.Value, out var y)) Properties.Y = y;
                    }

                    if (ext != null)
                    {
                        if (long.TryParse(ext.Attribute("cx")?.Value, out var cx)) Properties.Width = cx;
                        if (long.TryParse(ext.Attribute("cy")?.Value, out var cy)) Properties.Height = cy;
                    }

                    if (int.TryParse(xfrm.Attribute("rot")?.Value, out var rot))
                    {
                        Properties.Rotation = rot;
                    }
                }
            }

            public Dictionary<string, object?> ToDictionary()
            {
                var dict = new Dictionary<string, object?>
                {
                    ["type"] = Type,
                    ["id"] = Id
                };

                if (Properties != null)
                {
                    dict["x"] = Properties.X;
                    dict["y"] = Properties.Y;
                    dict["width"] = Properties.Width;
                    dict["height"] = Properties.Height;
                    if (Properties.Rotation.HasValue)
                        dict["rotation"] = Properties.Rotation;
                }

                return dict;
            }
        }

        /// <summary>
        /// Table element implementation.
        /// </summary>
        private sealed class TableElement : ITableElement
        {
            public string Type => "table";
            public string? Id { get; }
            public int RowCount { get; }
            public int ColumnCount { get; }

            public TableElement(XElement xmlElement, XNamespace a)
            {
                Id = xmlElement.Attribute("id")?.Value;

                var tblGrid = xmlElement.Element(a + "tblGrid");
                ColumnCount = tblGrid?.Elements(a + "gridCol").Count() ?? 0;

                var rows = xmlElement.Elements(a + "tr");
                RowCount = rows.Count();
            }

            public Dictionary<string, object?> ToDictionary()
            {
                return new Dictionary<string, object?>
                {
                    ["type"] = Type,
                    ["id"] = Id,
                    ["rows"] = RowCount,
                    ["columns"] = ColumnCount
                };
            }
        }

        /// <summary>
        /// Connector element implementation.
        /// </summary>
        private sealed class ConnectorElement : IConnectorElement
        {
            public string Type => "connector";
            public string? Id { get; }
            public string? StartConnectionId { get; }
            public string? EndConnectionId { get; }

            public ConnectorElement(XElement xmlElement, XNamespace p)
            {
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

                var nvCxnSpPr = xmlElement.Element(p + "nvCxnSpPr");
                var cNvPr = nvCxnSpPr?.Element(p + "cNvPr");
                Id = cNvPr?.Attribute("id")?.Value;

                var cNvCxnSpPr = nvCxnSpPr?.Element(p + "cNvCxnSpPr");
                if (cNvCxnSpPr != null)
                {
                    var stCxn = cNvCxnSpPr.Element(a + "stCxn");
                    var endCxn = cNvCxnSpPr.Element(a + "endCxn");

                    StartConnectionId = stCxn?.Attribute("id")?.Value;
                    EndConnectionId = endCxn?.Attribute("id")?.Value;
                }
            }

            public Dictionary<string, object?> ToDictionary()
            {
                return new Dictionary<string, object?>
                {
                    ["type"] = Type,
                    ["id"] = Id,
                    ["startConnectionId"] = StartConnectionId,
                    ["endConnectionId"] = EndConnectionId
                };
            }
        }
    }
}
