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

public class OpenXmlProvider : IOpenXmlProvider
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

        _logger.LogDebug("Opening package: {Path} with settings: FileMode={FileMode}, FileAccess={FileAccess}",
            path, settings.FileMode, settings.FileAccess);

        try
        {
            return await Task.Run(() =>
                Package.Open(path, settings.FileMode, settings.FileAccess, settings.FileShare),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error opening package: {Path}", path);
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
            _logger.LogError(ex, "Error getting package part: {PartUri}", partUri);
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
            _logger.LogError(ex, "Error getting relationships for part: {Uri}", part.Uri);
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
            _logger.LogError(ex, "Error getting stream for part: {Uri}", part.Uri);
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
            _logger.LogError(ex, "Error reading content from part: {Uri}", part.Uri);
            throw;
        }
    }
}
