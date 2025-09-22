using System;
using System.ComponentModel.DataAnnotations;
using System.Threading;
using System.Threading.Tasks;

namespace SharedXmlToJsonl.Providers;

public interface IConfigurationProvider
{
    T GetConfiguration<T>(string sectionName) where T : class, new();

    Task<T> GetConfigurationAsync<T>(
        string sectionName,
        CancellationToken cancellationToken = default) where T : class, new();

    ValidationResult ValidateConfiguration<T>(T configuration) where T : class;

    void BindConfiguration<T>(T configuration, string sectionName) where T : class;
}