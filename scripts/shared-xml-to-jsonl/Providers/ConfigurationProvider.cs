using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Providers;

public class ConfigurationProvider : IConfigurationProvider
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<ConfigurationProvider> _logger;

    public ConfigurationProvider(
        IConfiguration configuration,
        ILogger<ConfigurationProvider> logger)
    {
        _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public T GetConfiguration<T>(string sectionName) where T : class, new()
    {
        if (string.IsNullOrEmpty(sectionName))
            throw new ArgumentNullException(nameof(sectionName));

        _logger.LogDebug("Getting configuration for section: {SectionName}", sectionName);

        var configuration = new T();
        var section = _configuration.GetSection(sectionName);

        if (!section.Exists())
        {
            _logger.LogWarning("Configuration section not found: {SectionName}, using defaults", sectionName);
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
        if (configuration == null)
            throw new ArgumentNullException(nameof(configuration));

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
            _logger.LogError("Configuration validation failed: {Errors}", errors);
            return new ValidationResult(errors);
        }

        return ValidationResult.Success!;
    }

    public void BindConfiguration<T>(T configuration, string sectionName) where T : class
    {
        if (configuration == null)
            throw new ArgumentNullException(nameof(configuration));

        if (string.IsNullOrEmpty(sectionName))
            throw new ArgumentNullException(nameof(sectionName));

        var section = _configuration.GetSection(sectionName);
        if (section.Exists())
        {
            section.Bind(configuration);
            _logger.LogDebug("Bound configuration for section: {SectionName}", sectionName);
        }
        else
        {
            _logger.LogWarning("Configuration section not found: {SectionName}", sectionName);
        }
    }
}