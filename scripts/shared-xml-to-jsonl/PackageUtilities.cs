// File: PackageUtilities.cs
// Specification: Common utilities for handling Office Open XML package operations

using System;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;

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
}