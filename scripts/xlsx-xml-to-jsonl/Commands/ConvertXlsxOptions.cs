using System;
using System.ComponentModel.DataAnnotations;
using System.IO;
using SharedXmlToJsonl.Commands;

namespace XlsxXmlToJsonl.Commands
{
    /// <summary>
    /// Options for XLSX to JSONL conversion.
    /// </summary>
    public class ConvertXlsxOptions : CommandHandlerOptions
    {
        /// <summary>
        /// Gets or sets the maximum number of sheets to process.
        /// </summary>
        [Range(1, 1000, ErrorMessage = "MaxSheets must be between 1 and 1000")]
        public int MaxSheets { get; set; } = 1000;

        /// <summary>
        /// Gets or sets whether to include hidden sheets.
        /// </summary>
        public bool IncludeHiddenSheets { get; set; }

        /// <summary>
        /// Gets or sets whether to extract formulas.
        /// </summary>
        public bool ExtractFormulas { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to extract cell values.
        /// </summary>
        public bool ExtractValues { get; set; } = true;

        /// <summary>
        /// Gets or sets the start sheet index (0-based).
        /// </summary>
        [Range(0, 999, ErrorMessage = "StartSheetIndex must be between 0 and 999")]
        public int StartSheetIndex { get; set; } = 0;

        /// <summary>
        /// Gets or sets the end sheet index (0-based, -1 for all).
        /// </summary>
        [Range(-1, 999, ErrorMessage = "EndSheetIndex must be between -1 and 999")]
        public int EndSheetIndex { get; set; } = -1;

        /// <summary>
        /// Validates the XLSX-specific options.
        /// </summary>
        public override SharedXmlToJsonl.Commands.ValidationResult Validate()
        {
            var result = base.Validate();

            if (!File.Exists(InputPath))
            {
                result.Errors.Add($"Input file not found: {InputPath}");
            }
            else if (!InputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                result.Errors.Add("Input file must be a .xlsx file");
            }

            if (EndSheetIndex != -1 && EndSheetIndex < StartSheetIndex)
            {
                result.Errors.Add("EndSheetIndex must be greater than or equal to StartSheetIndex");
            }

            return result;
        }
    }
}
