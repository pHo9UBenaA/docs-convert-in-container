using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Providers;

public partial class ConfigurationProvider : IConfigurationProvider
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<ConfigurationProvider> _logger;

    public ConfigurationProvider(
        IConfiguration configuration,
        ILogger<ConfigurationProvider> logger)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        _configuration = configuration;
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;
    }

    public T GetConfiguration<T>(string sectionName) where T : class, new()
    {
        if (string.IsNullOrEmpty(sectionName))
            throw new ArgumentNullException(nameof(sectionName));

        LogGettingConfigurationForSection(_logger, sectionName);

        var configuration = new T();
        var section = _configuration.GetSection(sectionName);

        if (!section.Exists())
        {
            LogConfigurationSectionNotFound(_logger, sectionName);
            return configuration;
        }

        section.Bind(configuration);

        var validationResult = ValidateConfiguration(configuration);
        if (validationResult != ValidationResult.Success)
        {
            var errors = string.Join(", ", validationResult.ErrorMessage ?? "Validation failed");
            throw new InvalidOperationException($"Configuration validation failed for {sectionName}: {errors}");
        }

        return configuration;
    }

    public async Task<T> GetConfigurationAsync<T>(
        string sectionName,
        CancellationToken cancellationToken = default) where T : class, new()
    {
        return await Task.Run(() => GetConfiguration<T>(sectionName), cancellationToken).ConfigureAwait(false);
    }

    public ValidationResult ValidateConfiguration<T>(T configuration) where T : class
    {
        ArgumentNullException.ThrowIfNull(configuration);

        var validationContext = new ValidationContext(configuration);
        var validationResults = new List<ValidationResult>();

        bool isValid = Validator.TryValidateObject(
            configuration,
            validationContext,
            validationResults,
            validateAllProperties: true);

        if (!isValid)
        {
            var errors = string.Join("; ", validationResults.Select(r => r.ErrorMessage));
            LogConfigurationValidationFailed(_logger, errors);
            return new ValidationResult(errors);
        }

        return ValidationResult.Success!;
    }

    [LoggerMessage(
        EventId = 1001,
        Level = LogLevel.Debug,
        Message = "Getting configuration for section: {sectionName}")]
    private static partial void LogGettingConfigurationForSection(
        ILogger logger, string sectionName);

    [LoggerMessage(
        EventId = 1002,
        Level = LogLevel.Warning,
        Message = "Configuration section not found: {sectionName}, using defaults")]
    private static partial void LogConfigurationSectionNotFound(
        ILogger logger, string sectionName);

    [LoggerMessage(
        EventId = 1003,
        Level = LogLevel.Error,
        Message = "Configuration validation failed: {errors}")]
    private static partial void LogConfigurationValidationFailed(
        ILogger logger, string errors);

    [LoggerMessage(
        EventId = 1004,
        Level = LogLevel.Debug,
        Message = "Bound configuration for section: {sectionName}")]
    private static partial void LogBoundConfigurationForSection(
        ILogger logger, string sectionName);

    public void BindConfiguration<T>(T configuration, string sectionName) where T : class
    {
        ArgumentNullException.ThrowIfNull(configuration);

        if (string.IsNullOrEmpty(sectionName))
            throw new ArgumentNullException(nameof(sectionName));

        var section = _configuration.GetSection(sectionName);
        if (section.Exists())
        {
            section.Bind(configuration);
            LogBoundConfigurationForSection(_logger, sectionName);
        }
        else
        {
            LogConfigurationSectionNotFound(_logger, sectionName);
        }
    }
}
