using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;

namespace SharedXmlToJsonl.Commands
{
    /// <summary>
    /// Base class for command handler options with common properties and validation.
    /// </summary>
    public abstract class CommandHandlerOptions
    {
        /// <summary>
        /// Gets or sets the input file path.
        /// </summary>
        [Required(ErrorMessage = "Input path is required")]
        public string InputPath { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the output directory path.
        /// </summary>
        [Required(ErrorMessage = "Output directory is required")]
        public string OutputDirectory { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets whether to enable verbose output.
        /// </summary>
        public bool Verbose { get; set; }

        /// <summary>
        /// Validates the options and returns a validation result.
        /// </summary>
        /// <returns>The validation result containing any errors.</returns>
        [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Validation is only used for command-line options which are known at compile time")]
        public virtual ValidationResult Validate()
        {
            var result = new ValidationResult();

            // Validate using data annotations
            var context = new ValidationContext(this);
            var validationResults = new List<System.ComponentModel.DataAnnotations.ValidationResult>();

            if (!Validator.TryValidateObject(this, context, validationResults, validateAllProperties: true))
            {
                foreach (var validationResult in validationResults)
                {
                    result.Errors.Add(validationResult.ErrorMessage ?? "Validation error");
                }
            }

            // Check if output directory exists, create if it doesn't
            if (!string.IsNullOrWhiteSpace(OutputDirectory) && !Directory.Exists(OutputDirectory))
            {
                try
                {
                    Directory.CreateDirectory(OutputDirectory);
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"Failed to create output directory: {ex.Message}");
                }
            }

            return result;
        }
    }

    /// <summary>
    /// Represents the result of a validation operation.
    /// </summary>
    public class ValidationResult
    {
        /// <summary>
        /// Gets the list of validation errors.
        /// </summary>
        public List<string> Errors { get; } = new List<string>();

        /// <summary>
        /// Gets a value indicating whether the validation passed (no errors).
        /// </summary>
        public bool IsValid => Errors.Count == 0;
    }
}
