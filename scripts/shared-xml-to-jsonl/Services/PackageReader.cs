using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Services;

public interface IPackageReader
{
    Task<Package> OpenPackageAsync(string path, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<PackagePart>> GetPartsAsync(Package package, CancellationToken cancellationToken = default);
    Task<PackagePart?> GetPartByTypeAsync(Package package, string contentType, CancellationToken cancellationToken = default);
}

public partial class PackageReader : IPackageReader
{
    private readonly ILogger<PackageReader> _logger;

    public PackageReader(ILogger<PackageReader> logger)
    {
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;
    }

    public async Task<Package> OpenPackageAsync(
        string path,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}", path);

        LogOpeningPackage(_logger, path);

        try
        {
            return await Task.Run(() =>
                Package.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogErrorOpeningPackage(_logger, ex, path);
            throw;
        }
    }

    public async Task<IReadOnlyList<PackagePart>> GetPartsAsync(
        Package package,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(package);

        try
        {
            return await Task.Run(() =>
                package.GetParts().ToList(),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogErrorGettingPackageParts(_logger, ex);
            throw;
        }
    }

    public async Task<PackagePart?> GetPartByTypeAsync(
        Package package,
        string contentType,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(package);

        if (string.IsNullOrEmpty(contentType))
            throw new ArgumentNullException(nameof(contentType));

        try
        {
            return await Task.Run(() =>
                package.GetParts().FirstOrDefault(p => p.ContentType == contentType),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogErrorGettingPackagePartByType(_logger, ex, contentType);
            throw;
        }
    }

    [LoggerMessage(
        EventId = 2001,
        Level = LogLevel.Debug,
        Message = "Opening package: {path}")]
    private static partial void LogOpeningPackage(
        ILogger logger, string path);

    [LoggerMessage(
        EventId = 2002,
        Level = LogLevel.Error,
        Message = "Error opening package: {path}")]
    private static partial void LogErrorOpeningPackage(
        ILogger logger, Exception ex, string path);

    [LoggerMessage(
        EventId = 2003,
        Level = LogLevel.Error,
        Message = "Error getting package parts")]
    private static partial void LogErrorGettingPackageParts(
        ILogger logger, Exception ex);

    [LoggerMessage(
        EventId = 2004,
        Level = LogLevel.Error,
        Message = "Error getting package part by type: {contentType}")]
    private static partial void LogErrorGettingPackagePartByType(
        ILogger logger, Exception ex, string contentType);
}
