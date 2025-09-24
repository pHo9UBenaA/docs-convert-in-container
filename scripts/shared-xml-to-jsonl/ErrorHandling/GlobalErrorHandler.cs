using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using SharedXmlToJsonl.Exceptions;
using SharedXmlToJsonl.Resources;

namespace SharedXmlToJsonl.ErrorHandling;

public class GlobalErrorHandler : IGlobalErrorHandler
{
    private readonly ILogger<GlobalErrorHandler> _logger;
    private readonly IResourceService? _resourceService;

    public GlobalErrorHandler(
        ILogger<GlobalErrorHandler> logger,
        IResourceService? resourceService = null)
    {
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;
        _resourceService = resourceService;
    }

    public async Task<int> HandleErrorAsync(Exception exception, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(exception);

        await Task.CompletedTask; // Ensure async context

        switch (exception)
        {
            case DocumentProcessingException docEx:
                _logger.LogError(docEx,
                    "Document processing failed at {Stage} for {Path}. Page: {Page}, Element: {Element}",
                    docEx.Stage,
                    docEx.DocumentPath,
                    docEx.PageNumber ?? -1,
                    docEx.ElementId ?? "N/A");
                return CommonBase.ExitProcessingError;

            case ValidationException valEx:
                _logger.LogError(valEx,
                    "Validation failed with {Count} errors: {Message}",
                    valEx.ValidationErrors.Count,
                    valEx.Message);

                foreach (var error in valEx.ValidationErrors)
                {
                    _logger.LogError("Validation Error: {Error}", error);
                }
                return CommonBase.ExitUsageError;

            case OperationCanceledException opEx:
                _logger.LogWarning(opEx, "Operation was cancelled");
                return CommonBase.ExitProcessingError;

            case UnauthorizedAccessException uaEx:
                _logger.LogError(uaEx, "Access denied: {Message}", uaEx.Message);
                return CommonBase.ExitUsageError;

            case ArgumentException argEx:
                _logger.LogError(argEx, "Invalid argument: {Message}", argEx.Message);
                return CommonBase.ExitUsageError;

            case NotSupportedException nsEx:
                _logger.LogError(nsEx, "Operation not supported: {Message}", nsEx.Message);
                return CommonBase.ExitUsageError;

            case TimeoutException toEx:
                _logger.LogError(toEx, "Operation timed out: {Message}", toEx.Message);
                return CommonBase.ExitProcessingError;

            default:
                _logger.LogError(exception, "Unhandled exception: {Type} - {Message}",
                    exception.GetType().Name,
                    exception.Message);
                return CommonBase.ExitProcessingError;
        }
    }

    public void LogError(Exception exception, string context)
    {
        ArgumentNullException.ThrowIfNull(exception);

        _logger.LogError(exception, "Error in {Context}: {Message}",
            context ?? "Unknown",
            exception.Message);

        // Log inner exceptions
        var innerEx = exception.InnerException;
        var depth = 1;
        while (innerEx != null && depth <= 5) // Limit depth to prevent stack overflow
        {
            _logger.LogError(innerEx, "Inner Exception {Depth}: {Message}",
                depth,
                innerEx.Message);
            innerEx = innerEx.InnerException;
            depth++;
        }
    }

    public string FormatErrorMessage(Exception exception)
    {
        if (exception == null)
            return "Unknown error occurred";

        var message = exception switch
        {
            DocumentProcessingException docEx =>
                $"Failed to process document '{docEx.DocumentPath}' at stage {docEx.Stage}",

            ValidationException valEx =>
                $"Validation failed: {string.Join(", ", valEx.ValidationErrors)}",

            OperationCanceledException =>
                "Operation was cancelled by user",

            UnauthorizedAccessException =>
                $"Access denied: {exception.Message}",

            ArgumentException =>
                $"Invalid argument: {exception.Message}",

            NotSupportedException =>
                $"Not supported: {exception.Message}",

            TimeoutException =>
                "Operation timed out",

            _ => $"Error: {exception.Message}"
        };

        return message;
    }
}
