using System;
using System.Collections.Generic;
using System.Linq;

namespace SharedXmlToJsonl.Exceptions;

public class ValidationException : Exception
{
    public IReadOnlyList<string> ValidationErrors { get; }

    public ValidationException(string message)
        : base(message)
    {
        ValidationErrors = Array.Empty<string>();
    }

    public ValidationException(string message, IEnumerable<string> errors)
        : base(message)
    {
        ValidationErrors = errors?.ToList() ?? new List<string>();
    }

    public ValidationException(IEnumerable<string> errors)
        : base(FormatMessage(errors))
    {
        ValidationErrors = errors?.ToList() ?? new List<string>();
    }

    public ValidationException(string message, Exception innerException)
        : base(message, innerException)
    {
        ValidationErrors = Array.Empty<string>();
    }

    private static string FormatMessage(IEnumerable<string> errors)
    {
        var errorList = errors?.ToList() ?? new List<string>();
        if (!errorList.Any())
            return "Validation failed";

        return $"Validation failed with {errorList.Count} error(s): {string.Join("; ", errorList)}";
    }
}
