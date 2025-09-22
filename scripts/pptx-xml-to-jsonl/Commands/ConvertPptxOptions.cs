using System;
using System.ComponentModel.DataAnnotations;
using System.IO;
using SharedXmlToJsonl.Commands;

namespace PptxXmlToJsonl.Commands
{
    /// <summary>
    /// Options for PPTX to JSONL conversion.
    /// </summary>
    public class ConvertPptxOptions : CommandHandlerOptions
    {
        /// <summary>
        /// Gets or sets the maximum number of slides to process.
        /// </summary>
        [Range(1, 1000, ErrorMessage = "MaxSlides must be between 1 and 1000")]
        public int MaxSlides { get; set; } = 1000;

        /// <summary>
        /// Gets or sets whether to include hidden slides.
        /// </summary>
        public bool IncludeHiddenSlides { get; set; }

        /// <summary>
        /// Gets or sets whether to extract shape information.
        /// </summary>
        public bool ExtractShapes { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to extract text content.
        /// </summary>
        public bool ExtractText { get; set; } = true;

        /// <summary>
        /// Validates the PPTX-specific options.
        /// </summary>
        public override SharedXmlToJsonl.Commands.ValidationResult Validate()
        {
            var result = base.Validate();

            if (!File.Exists(InputPath))
            {
                result.Errors.Add($"Input file not found: {InputPath}");
            }
            else if (!InputPath.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
            {
                result.Errors.Add("Input file must be a .pptx file");
            }

            return result;
        }
    }
}