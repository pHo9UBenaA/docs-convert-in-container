using System;
using System.Globalization;
using System.Resources;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Resources;

public partial class ResourceService : IResourceService
{
    private readonly ResourceManager _errorMessages;
    private readonly ResourceManager _logMessages;
    private readonly ILogger<ResourceService> _logger;

    [LoggerMessage(EventId = 1, Level = LogLevel.Error, Message = "Error retrieving error message for key: {Key}")]
    private static partial void LogErrorRetrievingErrorMessage(ILogger logger, Exception exception, string key);

    [LoggerMessage(EventId = 2, Level = LogLevel.Error, Message = "Error retrieving log message for key: {Key}")]
    private static partial void LogErrorRetrievingLogMessage(ILogger logger, Exception exception, string key);

    [LoggerMessage(EventId = 3, Level = LogLevel.Error, Message = "Error retrieving string from {Resource} for key: {Key}")]
    private static partial void LogErrorRetrievingString(ILogger logger, Exception exception, string resource, string key);

    [LoggerMessage(EventId = 4, Level = LogLevel.Error, Message = "Error trying to get string from {Resource} for key: {Key}")]
    private static partial void LogErrorTryingToGetString(ILogger logger, Exception exception, string resource, string key);

    public ResourceService(ILogger<ResourceService> logger)
    {
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;

        _errorMessages = new ResourceManager(
            "SharedXmlToJsonl.Resources.ErrorMessages",
            typeof(ResourceService).Assembly);

        _logMessages = new ResourceManager(
            "SharedXmlToJsonl.Resources.LogMessages",
            typeof(ResourceService).Assembly);
    }

    public string GetErrorMessage(string key, params object[] args)
    {
        if (string.IsNullOrEmpty(key))
            throw new ArgumentNullException(nameof(key));

        try
        {
            var template = _errorMessages.GetString(key, CultureInfo.CurrentCulture) ?? key;
            return args != null && args.Length > 0
                ? string.Format(CultureInfo.CurrentCulture, template, args)
                : template;
        }
        catch (Exception ex)
        {
            LogErrorRetrievingErrorMessage(_logger, ex, key);
            return key;
        }
    }

    public string GetLogMessage(string key, params object[] args)
    {
        if (string.IsNullOrEmpty(key))
            throw new ArgumentNullException(nameof(key));

        try
        {
            var template = _logMessages.GetString(key, CultureInfo.CurrentCulture) ?? key;
            return args != null && args.Length > 0
                ? string.Format(CultureInfo.CurrentCulture, template, args)
                : template;
        }
        catch (Exception ex)
        {
            LogErrorRetrievingLogMessage(_logger, ex, key);
            return key;
        }
    }

    public string GetString(string resourceName, string key, params object[] args)
    {
        if (string.IsNullOrEmpty(resourceName))
            throw new ArgumentNullException(nameof(resourceName));

        if (string.IsNullOrEmpty(key))
            throw new ArgumentNullException(nameof(key));

        try
        {
            var resourceManager = new ResourceManager(
                $"SharedXmlToJsonl.Resources.{resourceName}",
                typeof(ResourceService).Assembly);

            var template = resourceManager.GetString(key, CultureInfo.CurrentCulture) ?? key;
            return args != null && args.Length > 0
                ? string.Format(CultureInfo.CurrentCulture, template, args)
                : template;
        }
        catch (Exception ex)
        {
            LogErrorRetrievingString(_logger, ex, resourceName, key);
            return key;
        }
    }

    public bool TryGetString(string resourceName, string key, out string? value)
    {
        value = null;

        if (string.IsNullOrEmpty(resourceName) || string.IsNullOrEmpty(key))
            return false;

        try
        {
            var resourceManager = new ResourceManager(
                $"SharedXmlToJsonl.Resources.{resourceName}",
                typeof(ResourceService).Assembly);

            value = resourceManager.GetString(key, CultureInfo.CurrentCulture);
            return value != null;
        }
        catch (Exception ex)
        {
            LogErrorTryingToGetString(_logger, ex, resourceName, key);
            return false;
        }
    }
}
