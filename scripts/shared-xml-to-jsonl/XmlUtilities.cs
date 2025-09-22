using System.Xml.Linq;

namespace SharedXmlToJsonl;

/// <summary>
/// Utilities for parsing OpenXML/DrawingML elements
/// </summary>
public static class XmlUtilities
{
    /// <summary>
    /// Extract transform information from xfrm element
    /// </summary>
    public static Transform? ExtractTransformFromXfrm(XElement? xfrm, XNamespace a)
    {
        if (xfrm == null) return null;

        Position? position = null;
        Size? size = null;
        double? rotation = null;

        // Extract offset (position)
        var off = xfrm.Element(a + "off");
        if (off != null)
        {
            if (long.TryParse(off.Attribute("x")?.Value, out var x) &&
                long.TryParse(off.Attribute("y")?.Value, out var y))
            {
                position = new Position(x, y);
            }
        }

        // Extract extents (size)
        var ext = xfrm.Element(a + "ext");
        if (ext != null)
        {
            if (long.TryParse(ext.Attribute("cx")?.Value, out var cx) &&
                long.TryParse(ext.Attribute("cy")?.Value, out var cy))
            {
                size = new Size(cx, cy);
            }
        }

        // Extract rotation
        var rot = xfrm.Attribute("rot")?.Value;
        if (!string.IsNullOrEmpty(rot) && long.TryParse(rot, out var rotValue))
        {
            // Convert from 60000ths of a degree to degrees
            rotation = rotValue / 60000.0;
        }

        return new Transform(position, size, rotation);
    }

    /// <summary>
    /// Extract custom geometry from custGeom element
    /// </summary>
    public static CustomGeometry? ExtractCustomGeometry(XElement? custGeom, XNamespace a)
    {
        if (custGeom == null) return null;

        var pathLst = custGeom.Element(a + "pathLst");
        if (pathLst == null) return null;

        var paths = new List<string>();
        foreach (var path in pathLst.Elements(a + "path"))
        {
            var pathData = new List<string>();

            // Process move to commands
            foreach (var moveTo in path.Elements(a + "moveTo"))
            {
                var pt = moveTo.Element(a + "pt");
                if (pt != null)
                {
                    var x = pt.Attribute("x")?.Value;
                    var y = pt.Attribute("y")?.Value;
                    if (x != null && y != null)
                    {
                        pathData.Add($"M {x} {y}");
                    }
                }
            }

            // Process line to commands
            foreach (var lineTo in path.Elements(a + "lnTo"))
            {
                var pt = lineTo.Element(a + "pt");
                if (pt != null)
                {
                    var x = pt.Attribute("x")?.Value;
                    var y = pt.Attribute("y")?.Value;
                    if (x != null && y != null)
                    {
                        pathData.Add($"L {x} {y}");
                    }
                }
            }

            // Process cubic bezier commands
            foreach (var cubicBez in path.Elements(a + "cubicBezTo"))
            {
                var pts = cubicBez.Elements(a + "pt").ToList();
                if (pts.Count == 3)
                {
                    var coords = pts.Select(pt => $"{pt.Attribute("x")?.Value} {pt.Attribute("y")?.Value}");
                    pathData.Add($"C {string.Join(" ", coords)}");
                }
            }

            // Process close commands
            if (path.Elements(a + "close").Any())
            {
                pathData.Add("Z");
            }

            if (pathData.Count > 0)
            {
                paths.Add(string.Join(" ", pathData));
            }
        }

        if (paths.Count == 0) return null;

        var fillRule = pathLst.Attribute("fill")?.Value;
        return new CustomGeometry(string.Join(" ", paths), fillRule);
    }

    /// <summary>
    /// Extract table cells from DrawingML table
    /// </summary>
    public static List<TableCell> ExtractTableCells(XElement tbl, XNamespace a)
    {
        var cells = new List<TableCell>();
        var rows = tbl.Elements(a + "tr").ToList();

        for (int rowIdx = 0; rowIdx < rows.Count; rowIdx++)
        {
            var row = rows[rowIdx];
            var tcElements = row.Elements(a + "tc").ToList();

            for (int colIdx = 0; colIdx < tcElements.Count; colIdx++)
            {
                var tc = tcElements[colIdx];

                // Extract text from cell
                var txBody = tc.Element(a + "txBody");
                var cellTexts = new List<string>();
                if (txBody != null)
                {
                    foreach (var p in txBody.Elements(a + "p"))
                    {
                        var runs = p.Elements(a + "r");
                        foreach (var r in runs)
                        {
                            var t = r.Element(a + "t");
                            if (t != null)
                            {
                                cellTexts.Add(t.Value);
                            }
                        }
                    }
                }

                // Extract span information
                int? rowSpan = null;
                int? colSpan = null;
                var gridSpan = tc.Attribute("gridSpan")?.Value;
                if (int.TryParse(gridSpan, out var span))
                {
                    colSpan = span;
                }
                var rowSpanAttr = tc.Attribute("rowSpan")?.Value;
                if (int.TryParse(rowSpanAttr, out var rSpan))
                {
                    rowSpan = rSpan;
                }

                cells.Add(new TableCell(
                    rowIdx,
                    colIdx,
                    cellTexts.Count > 0 ? string.Join(" ", cellTexts) : null,
                    rowSpan,
                    colSpan
                ));
            }
        }

        return cells;
    }

    /// <summary>
    /// Extract cell reference from Excel drawing anchor
    /// </summary>
    public static (string Cell, int Col, int Row)? ExtractCellReference(XElement? cellElement, XNamespace xdr)
    {
        if (cellElement == null) return null;

        var col = cellElement.Element(xdr + "col")?.Value;
        var row = cellElement.Element(xdr + "row")?.Value;

        if (!string.IsNullOrEmpty(col) && !string.IsNullOrEmpty(row) &&
            int.TryParse(col, out var colNum) && int.TryParse(row, out var rowNum))
        {
            // Convert column number to letter(s) (0=A, 1=B, etc.)
            var colLetter = GetColumnLetter(colNum);
            var cellRef = $"{colLetter}{rowNum + 1}"; // Rows are 0-indexed in XML but 1-indexed in Excel

            return (cellRef, colNum, rowNum);
        }

        return null;
    }

    /// <summary>
    /// Convert column index to Excel column letter(s)
    /// </summary>
    private static string GetColumnLetter(int columnIndex)
    {
        var columnLetter = "";
        while (columnIndex >= 0)
        {
            columnLetter = (char)('A' + columnIndex % 26) + columnLetter;
            columnIndex = columnIndex / 26 - 1;
        }
        return columnLetter;
    }

    /// <summary>
    /// Gets the position from a transform element
    /// </summary>
    public static Position? GetPositionFromTransform(XElement? xfrm, XNamespace a)
    {
        if (xfrm == null) return null;

        var off = xfrm.Element(a + "off");
        if (off != null)
        {
            if (long.TryParse(off.Attribute("x")?.Value, out var x) &&
                long.TryParse(off.Attribute("y")?.Value, out var y))
            {
                return new Position(x, y);
            }
        }

        return null;
    }

    /// <summary>
    /// Gets the size from a transform element
    /// </summary>
    public static Size? GetSizeFromTransform(XElement? xfrm, XNamespace a)
    {
        if (xfrm == null) return null;

        var ext = xfrm.Element(a + "ext");
        if (ext != null)
        {
            if (long.TryParse(ext.Attribute("cx")?.Value, out var cx) &&
                long.TryParse(ext.Attribute("cy")?.Value, out var cy))
            {
                return new Size(cx, cy);
            }
        }

        return null;
    }
}