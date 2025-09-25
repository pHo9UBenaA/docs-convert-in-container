using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Providers;

public partial class OpenXmlProvider : IOpenXmlProvider
{
    private readonly ILogger<OpenXmlProvider> _logger;

    public OpenXmlProvider(ILogger<OpenXmlProvider> logger)
    {
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;
    }

    public async Task<Package> OpenPackageAsync(
        string path,
        OpenSettings settings,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        ArgumentNullException.ThrowIfNull(settings);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Package file not found: {path}", path);

        LogOpeningPackageWithSettings(_logger, path, settings.FileMode.ToString(), settings.FileAccess.ToString());

        try
        {
            return await Task.Run(() =>
                Package.Open(path, settings.FileMode, settings.FileAccess, settings.FileShare),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogErrorOpeningPackage(_logger, ex, path);
            throw;
        }
    }

    public async Task<PackagePart?> GetPartAsync(
        Package package,
        string partUri,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(package);

        if (string.IsNullOrEmpty(partUri))
            throw new ArgumentNullException(nameof(partUri));

        try
        {
            return await Task.Run(() =>
            {
                var uri = new Uri(partUri, UriKind.Relative);
                return package.PartExists(uri) ? package.GetPart(uri) : null;
            }, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogErrorGettingPackagePart(_logger, ex, partUri);
            throw;
        }
    }

    public async Task<IReadOnlyList<PackageRelationship>> GetRelationshipsAsync(
        PackagePart part,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(part);

        try
        {
            return await Task.Run(() =>
                part.GetRelationships().ToList(),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogErrorGettingRelationships(_logger, ex, part.Uri.ToString());
            throw;
        }
    }

    public async Task<Stream> GetPartStreamAsync(
        PackagePart part,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(part);

        try
        {
            return await Task.Run(() => part.GetStream(FileMode.Open, FileAccess.Read),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogErrorGettingStreamForPart(_logger, ex, part.Uri.ToString());
            throw;
        }
    }

    public async Task<string> GetPartContentAsync(
        PackagePart part,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(part);

        try
        {
            await using var stream = await GetPartStreamAsync(part, cancellationToken).ConfigureAwait(false);
            using var reader = new StreamReader(stream, Encoding.UTF8);
            return await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogErrorReadingContentFromPart(_logger, ex, part.Uri.ToString());
            throw;
        }
    }

    [LoggerMessage(
        EventId = 6001,
        Level = LogLevel.Debug,
        Message = "Opening package: {path} with settings: FileMode={fileMode}, FileAccess={fileAccess}")]
    private static partial void LogOpeningPackageWithSettings(
        ILogger logger, string path, string fileMode, string fileAccess);

    [LoggerMessage(
        EventId = 6002,
        Level = LogLevel.Error,
        Message = "Error opening package: {path}")]
    private static partial void LogErrorOpeningPackage(
        ILogger logger, Exception ex, string path);

    [LoggerMessage(
        EventId = 6003,
        Level = LogLevel.Error,
        Message = "Error getting package part: {partUri}")]
    private static partial void LogErrorGettingPackagePart(
        ILogger logger, Exception ex, string partUri);

    [LoggerMessage(
        EventId = 6004,
        Level = LogLevel.Error,
        Message = "Error getting relationships for part: {uri}")]
    private static partial void LogErrorGettingRelationships(
        ILogger logger, Exception ex, string uri);

    [LoggerMessage(
        EventId = 6005,
        Level = LogLevel.Error,
        Message = "Error getting stream for part: {uri}")]
    private static partial void LogErrorGettingStreamForPart(
        ILogger logger, Exception ex, string uri);

    [LoggerMessage(
        EventId = 6006,
        Level = LogLevel.Error,
        Message = "Error reading content from part: {uri}")]
    private static partial void LogErrorReadingContentFromPart(
        ILogger logger, Exception ex, string uri);
}
