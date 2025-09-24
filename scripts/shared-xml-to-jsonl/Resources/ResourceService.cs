using System;
using System.Globalization;
using System.Resources;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Resources;

public class ResourceService : IResourceService
{
    private readonly ResourceManager _errorMessages;
    private readonly ResourceManager _logMessages;
    private readonly ILogger<ResourceService> _logger;

    public ResourceService(ILogger<ResourceService> logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));

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
            _logger.LogError(ex, "Error retrieving error message for key: {Key}", key);
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
            _logger.LogError(ex, "Error retrieving log message for key: {Key}", key);
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
            _logger.LogError(ex, "Error retrieving string from {Resource} for key: {Key}",
                resourceName, key);
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
            _logger.LogError(ex, "Error trying to get string from {Resource} for key: {Key}",
                resourceName, key);
            return false;
        }
    }
}
