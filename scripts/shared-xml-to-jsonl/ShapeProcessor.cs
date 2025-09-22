using System.Xml.Linq;

namespace SharedXmlToJsonl;

/// <summary>
/// Common shape processing utilities for PPTX and XLSX
/// </summary>
public static class ShapeProcessor
{
    /// <summary>
    /// Extract line properties from shape properties element
    /// </summary>
    public static LineProperties? ExtractLineProperties(XElement? spPr, XNamespace a)
    {
        if (spPr == null) return null;

        var ln = spPr.Element(a + "ln");
        if (ln == null) return null;

        string? color = null;
        long? width = null;
        string? dashStyle = null;
        string? compoundLineType = null;

        // Extract line width
        var widthAttr = ln.Attribute("w");
        if (widthAttr != null && long.TryParse(widthAttr.Value, out var w))
        {
            width = w;
        }

        // Extract line color
        var solidFill = ln.Element(a + "solidFill");
        if (solidFill != null)
        {
            var srgbClr = solidFill.Element(a + "srgbClr");
            if (srgbClr != null)
            {
                color = srgbClr.Attribute("val")?.Value;
            }
            else
            {
                var schemeClr = solidFill.Element(a + "schemeClr");
                if (schemeClr != null)
                {
                    color = schemeClr.Attribute("val")?.Value;
                }
            }
        }

        // Extract dash style
        var prstDash = ln.Element(a + "prstDash");
        if (prstDash != null)
        {
            dashStyle = prstDash.Attribute("val")?.Value;
        }

        // Extract compound line type
        var cmpd = ln.Attribute("cmpd");
        if (cmpd != null)
        {
            compoundLineType = cmpd.Value;
        }

        return new LineProperties(color, width, dashStyle, compoundLineType);
    }

    /// <summary>
    /// Extract fill information from shape properties
    /// </summary>
    public static (bool? HasFill, string? FillColor) ExtractFillInfo(XElement? spPr, XNamespace a)
    {
        if (spPr == null) return (null, null);

        // Check for no fill
        var noFill = spPr.Element(a + "noFill");
        if (noFill != null)
        {
            return (false, null);
        }

        // Check for solid fill
        var solidFill = spPr.Element(a + "solidFill");
        if (solidFill != null)
        {
            var srgbClr = solidFill.Element(a + "srgbClr");
            if (srgbClr != null)
            {
                var color = srgbClr.Attribute("val")?.Value;
                return (true, color);
            }

            var schemeClr = solidFill.Element(a + "schemeClr");
            if (schemeClr != null)
            {
                var color = schemeClr.Attribute("val")?.Value;
                return (true, color);
            }

            return (true, null);
        }

        // Check for gradient fill
        var gradFill = spPr.Element(a + "gradFill");
        if (gradFill != null)
        {
            return (true, "gradient");
        }

        // Check for pattern fill
        var pattFill = spPr.Element(a + "pattFill");
        if (pattFill != null)
        {
            return (true, "pattern");
        }

        // Check for picture fill
        var blipFill = spPr.Element(a + "blipFill");
        if (blipFill != null)
        {
            return (true, "picture");
        }

        return (null, null);
    }

    /// <summary>
    /// Extract shape transform from shape properties (common for PPTX)
    /// </summary>
    public static Transform? ExtractTransform(XElement? spPr, XNamespace a)
    {
        if (spPr == null) return null;
        var xfrm = spPr.Element(a + "xfrm");
        return XmlUtilities.ExtractTransformFromXfrm(xfrm, a);
    }

    /// <summary>
    /// Generate unique element ID for tracking
    /// </summary>
    public static string GenerateElementId(string prefix, int slideOrSheetNumber, int elementIndex)
    {
        return $"{prefix}_{slideOrSheetNumber}_{elementIndex}";
    }

    /// <summary>
    /// Extract text from text body element (common structure in PPTX/XLSX)
    /// </summary>
    public static List<string> ExtractTextFromTxBody(XElement? txBody, XNamespace a)
    {
        var texts = new List<string>();
        if (txBody == null) return texts;

        foreach (var p in txBody.Elements(a + "p"))
        {
            var paragraphTexts = new List<string>();
            foreach (var r in p.Elements(a + "r"))
            {
                var t = r.Element(a + "t");
                if (t != null && !string.IsNullOrEmpty(t.Value))
                {
                    paragraphTexts.Add(t.Value);
                }
            }

            if (paragraphTexts.Count > 0)
            {
                texts.Add(string.Join("", paragraphTexts));
            }
        }

        return texts;
    }

    /// <summary>
    /// Extract preset shape type or determine if custom
    /// </summary>
    public static string? ExtractShapeType(XElement? spPr, XNamespace a)
    {
        if (spPr == null) return null;

        var prstGeom = spPr.Element(a + "prstGeom");
        if (prstGeom != null)
        {
            return prstGeom.Attribute("prst")?.Value;
        }

        var custGeom = spPr.Element(a + "custGeom");
        if (custGeom != null)
        {
            return "custom";
        }

        return null;
    }

    /// <summary>
    /// Extract connector type from connector shape
    /// </summary>
    public static string? ExtractConnectorType(XElement? spPr, XNamespace a)
    {
        if (spPr == null) return null;

        var prstGeom = spPr.Element(a + "prstGeom");
        return prstGeom?.Attribute("prst")?.Value;
    }
}