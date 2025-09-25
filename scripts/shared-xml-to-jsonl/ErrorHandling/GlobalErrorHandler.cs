using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using SharedXmlToJsonl.Exceptions;
using SharedXmlToJsonl.Resources;

namespace SharedXmlToJsonl.ErrorHandling;

public partial class GlobalErrorHandler : IGlobalErrorHandler
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
                LogDocumentProcessingFailed(_logger, docEx, docEx.Stage.ToString(), docEx.DocumentPath, docEx.PageNumber ?? -1, docEx.ElementId ?? "N/A");
                return CommonBase.ExitProcessingError;

            case ValidationException valEx:
                LogValidationFailed(_logger, valEx, valEx.ValidationErrors.Count, valEx.Message ?? string.Empty);

                foreach (var error in valEx.ValidationErrors)
                {
                    LogValidationError(_logger, error);
                }
                return CommonBase.ExitUsageError;

            case OperationCanceledException opEx:
                LogOperationCancelled(_logger, opEx);
                return CommonBase.ExitProcessingError;

            case UnauthorizedAccessException uaEx:
                LogAccessDenied(_logger, uaEx, uaEx.Message ?? string.Empty);
                return CommonBase.ExitUsageError;

            case ArgumentException argEx:
                LogInvalidArgument(_logger, argEx, argEx.Message ?? string.Empty);
                return CommonBase.ExitUsageError;

            case NotSupportedException nsEx:
                LogOperationNotSupported(_logger, nsEx, nsEx.Message ?? string.Empty);
                return CommonBase.ExitUsageError;

            case TimeoutException toEx:
                LogOperationTimedOut(_logger, toEx, toEx.Message ?? string.Empty);
                return CommonBase.ExitProcessingError;

            default:
                LogUnhandledException(_logger, exception, exception.GetType().Name, exception.Message ?? string.Empty);
                return CommonBase.ExitProcessingError;
        }
    }

    public void LogError(Exception exception, string context)
    {
        ArgumentNullException.ThrowIfNull(exception);

        LogErrorInContext(_logger, exception, context ?? "Unknown", exception.Message ?? string.Empty);

        // Log inner exceptions
        var innerEx = exception.InnerException;
        var depth = 1;
        while (innerEx != null && depth <= 5) // Limit depth to prevent stack overflow
        {
            LogInnerException(_logger, innerEx, depth, innerEx.Message ?? string.Empty);
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

    [LoggerMessage(
        EventId = 9001,
        Level = LogLevel.Error,
        Message = "Document processing failed at {stage} for {path}. Page: {page}, Element: {element}")]
    private static partial void LogDocumentProcessingFailed(
        ILogger logger, Exception ex, string stage, string path, int page, string element);

    [LoggerMessage(
        EventId = 9002,
        Level = LogLevel.Error,
        Message = "Validation failed with {count} errors: {message}")]
    private static partial void LogValidationFailed(
        ILogger logger, Exception ex, int count, string message);

    [LoggerMessage(
        EventId = 9003,
        Level = LogLevel.Error,
        Message = "Validation Error: {error}")]
    private static partial void LogValidationError(
        ILogger logger, string error);

    [LoggerMessage(
        EventId = 9004,
        Level = LogLevel.Warning,
        Message = "Operation was cancelled")]
    private static partial void LogOperationCancelled(
        ILogger logger, Exception ex);

    [LoggerMessage(
        EventId = 9005,
        Level = LogLevel.Error,
        Message = "Access denied: {message}")]
    private static partial void LogAccessDenied(
        ILogger logger, Exception ex, string message);

    [LoggerMessage(
        EventId = 9006,
        Level = LogLevel.Error,
        Message = "Invalid argument: {message}")]
    private static partial void LogInvalidArgument(
        ILogger logger, Exception ex, string message);

    [LoggerMessage(
        EventId = 9007,
        Level = LogLevel.Error,
        Message = "Operation not supported: {message}")]
    private static partial void LogOperationNotSupported(
        ILogger logger, Exception ex, string message);

    [LoggerMessage(
        EventId = 9008,
        Level = LogLevel.Error,
        Message = "Operation timed out: {message}")]
    private static partial void LogOperationTimedOut(
        ILogger logger, Exception ex, string message);

    [LoggerMessage(
        EventId = 9009,
        Level = LogLevel.Error,
        Message = "Unhandled exception: {type} - {message}")]
    private static partial void LogUnhandledException(
        ILogger logger, Exception ex, string type, string message);

    [LoggerMessage(
        EventId = 9010,
        Level = LogLevel.Error,
        Message = "Error in {context}: {message}")]
    private static partial void LogErrorInContext(
        ILogger logger, Exception ex, string context, string message);

    [LoggerMessage(
        EventId = 9011,
        Level = LogLevel.Error,
        Message = "Inner Exception {depth}: {message}")]
    private static partial void LogInnerException(
        ILogger logger, Exception ex, int depth, string message);
}
