// File: PackageUtilities.cs
// Specification: Common utilities for handling Office Open XML package operations

using System;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace SharedXmlToJsonl;

/// <summary>
/// Common utilities for Office Open XML package operations
/// </summary>
public static class PackageUtilities
{
    /// <summary>
    /// Tries to get a package part from a package by its URI
    /// </summary>
    /// <param name="package">The package to search in</param>
    /// <param name="partName">The part name to find</param>
    /// <param name="packagePart">The found package part (if any)</param>
    /// <returns>True if the part was found, false otherwise</returns>
    public static bool TryGetPackagePart(Package package, string partName, out PackagePart packagePart)
    {
        packagePart = null!;
        try
        {
            var uri = PackUriHelper.CreatePartUri(new Uri(partName, UriKind.RelativeOrAbsolute));
            if (package.PartExists(uri))
            {
                packagePart = package.GetPart(uri);
                return true;
            }
        }
        catch (ArgumentException)
        {
            // Relationship parts such as "/_rels/.rels" or invalid URIs fall through here.
        }
        return false;
    }

    /// <summary>
    /// Determines the fallback content type for a part when not found in the package
    /// </summary>
    /// <param name="partName">The part name to determine content type for</param>
    /// <returns>The appropriate content type</returns>
    public static string DetermineFallbackContentType(string partName)
    {
        if (partName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
        {
            return NamespaceConstants.RelationshipContentType;
        }

        return NamespaceConstants.DefaultXmlContentType;
    }

    /// <summary>
    /// Determines if a ZIP entry is an XML entry that should be processed
    /// </summary>
    /// <param name="entry">The ZIP archive entry to check</param>
    /// <returns>True if the entry is an XML file, false otherwise</returns>
    public static bool IsXmlEntry(ZipArchiveEntry entry)
    {
        var name = entry.FullName;
        if (name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Extracts relationships from a package part
    /// </summary>
    /// <param name="packagePart">The package part to extract relationships from</param>
    /// <returns>An ordered collection of relationship information</returns>
    public static IReadOnlyList<RelationshipInfo> ExtractRelationships(PackagePart packagePart)
    {
        if (string.Equals(packagePart.ContentType, NamespaceConstants.RelationshipContentType, StringComparison.OrdinalIgnoreCase))
        {
            return Array.Empty<RelationshipInfo>();
        }

        var relationships = packagePart
            .GetRelationships()
            .Select(rel => new RelationshipInfo(
                rel.Id,
                rel.RelationshipType,
                rel.TargetUri?.ToString() ?? string.Empty,
                rel.TargetMode.ToString()))
            .OrderBy(rel => rel.Id, StringComparer.Ordinal)
            .ToArray();

        return relationships;
    }

    /// <summary>
    /// Gets the presentation part from a PowerPoint package
    /// </summary>
    public static PackagePart GetPresentationPart(Package package)
    {
        var presentationPart = package.GetParts()
            .FirstOrDefault(p => p.Uri.ToString().EndsWith("presentation.xml", StringComparison.OrdinalIgnoreCase));

        if (presentationPart == null)
            throw new InvalidOperationException("Presentation part not found in package");

        return presentationPart;
    }

    /// <summary>
    /// Gets an XDocument from a package part
    /// </summary>
    public static XDocument GetXDocument(PackagePart part)
    {
        using var stream = part.GetStream();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// Gets a related part by relationship ID
    /// </summary>
    public static PackagePart GetRelatedPart(Package package, PackagePart sourcePart, string relationshipId)
    {
        var relationship = sourcePart.GetRelationship(relationshipId);
        var targetUri = PackUriHelper.ResolvePartUri(sourcePart.Uri, relationship.TargetUri);
        return package.GetPart(targetUri);
    }

    /// <summary>
    /// Gets the workbook part from an Excel package
    /// </summary>
    public static PackagePart GetWorkbookPart(Package package)
    {
        var workbookPart = package.GetParts()
            .FirstOrDefault(p => p.Uri.ToString().EndsWith("workbook.xml", StringComparison.OrdinalIgnoreCase));

        if (workbookPart == null)
            throw new InvalidOperationException("Workbook part not found in package");

        return workbookPart;
    }

    /// <summary>
    /// Gets the shared strings part from an Excel package
    /// </summary>
    public static PackagePart? GetSharedStringsPart(Package package)
    {
        var sharedStringsPart = package.GetParts()
            .FirstOrDefault(p => p.Uri.ToString().EndsWith("sharedStrings.xml", StringComparison.OrdinalIgnoreCase));

        return sharedStringsPart;  // May be null if no shared strings exist
    }

    /// <summary>
    /// Gets the styles part from an Excel package
    /// </summary>
    public static PackagePart? GetStylesPart(Package package)
    {
        var stylesPart = package.GetParts()
            .FirstOrDefault(p => p.Uri.ToString().EndsWith("styles.xml", StringComparison.OrdinalIgnoreCase));

        return stylesPart;  // May be null if no styles exist
    }
}