using System;
using System.Globalization;
using System.IO;
using System.Linq;

namespace SharedXmlToJsonl
{
    /// <summary>
    /// Provides utilities for document processing.
    /// </summary>
    public static class DocumentUtilities
    {
        /// <summary>
        /// Extracts a document number from a part name based on prefix pattern.
        /// </summary>
        /// <param name="partName">The part name to extract from</param>
        /// <param name="prefix">The prefix pattern (e.g., "/ppt/slides/slide" or "/xl/worksheets/sheet")</param>
        /// <returns>The extracted document number, or null if extraction failed</returns>
        public static int? TryExtractDocumentNumber(string partName, string prefix)
        {
            if (string.IsNullOrEmpty(partName) || string.IsNullOrEmpty(prefix))
            {
                return null;
            }

            if (!partName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var suffix = partName.Substring(prefix.Length);
            var digits = new string(suffix.TakeWhile(char.IsDigit).ToArray());
            if (digits.Length == 0)
            {
                return null;
            }

            if (int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out var value))
            {
                return value;
            }

            return null;
        }

        /// <summary>
        /// Extracts slide number from a PPTX part name.
        /// </summary>
        /// <param name="partName">The part name to extract from</param>
        /// <returns>The slide number, or null if extraction failed</returns>
        public static int? TryExtractSlideNumber(string partName)
        {
            const string slidePrefix = "/ppt/slides/slide";
            return TryExtractDocumentNumber(partName, slidePrefix);
        }

        /// <summary>
        /// Extracts sheet number from an XLSX part name.
        /// </summary>
        /// <param name="partName">The part name to extract from</param>
        /// <returns>The sheet number, or null if extraction failed</returns>
        public static int? TryExtractSheetNumber(string partName)
        {
            const string sheetPrefix = "/xl/worksheets/sheet";
            return TryExtractDocumentNumber(partName, sheetPrefix);
        }

        /// <summary>
        /// Gets the document name from a file path.
        /// </summary>
        /// <param name="path">The file path</param>
        /// <returns>The document name without extension</returns>
        public static string GetDocumentNameFromPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return "unknown";

            return Path.GetFileNameWithoutExtension(path) ?? "unknown";
        }
    }
}