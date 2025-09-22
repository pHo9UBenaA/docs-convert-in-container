using System.Xml.Linq;

namespace SharedXmlToJsonl;

/// <summary>
/// Common XML namespace definitions for Office Open XML processing
/// </summary>
public static class NamespaceConstants
{
    // Core namespaces
    public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    // Spreadsheet namespaces
    public static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // Lowercase aliases for backward compatibility
    public static readonly XNamespace a = A;
    public static readonly XNamespace p = P;
    public static readonly XNamespace r = R;
    public static readonly XNamespace spreadsheet = S;

    // Additional drawing namespaces
    public static readonly XNamespace A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
    public static readonly XNamespace A15 = "http://schemas.microsoft.com/office/drawing/2012/main";
    public static readonly XNamespace A16 = "http://schemas.microsoft.com/office/drawing/2014/main";
    public static readonly XNamespace X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    // Content types
    public const string RelationshipContentType = "application/vnd.openxmlformats-package.relationships+xml";
    public const string DefaultXmlContentType = "application/xml";

    // Relationship types
    public const string SlideRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    public const string WorksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    public const string DrawingRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    public const string ChartRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    public const string ImageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    public const string OleObjectRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
}