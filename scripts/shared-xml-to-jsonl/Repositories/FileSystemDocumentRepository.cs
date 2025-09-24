using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Repositories;

public class FileSystemDocumentRepository : IDocumentRepository
{
    private readonly ILogger<FileSystemDocumentRepository> _logger;
    private const int DefaultBufferSize = 4096;

    public FileSystemDocumentRepository(ILogger<FileSystemDocumentRepository> logger)
    {
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;
    }

    public async Task<bool> ExistsAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        return await Task.Run(() => File.Exists(path), cancellationToken).ConfigureAwait(false);
    }

    public async Task<Stream> OpenReadStreamAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}", path);

        _logger.LogDebug("Opening read stream for: {Path}", path);

        try
        {
            return await Task.Run(() =>
                new FileStream(
                    path,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read,
                    DefaultBufferSize,
                    useAsync: true),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error opening read stream for: {Path}", path);
            throw;
        }
    }

    public async Task<Stream> OpenWriteStreamAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        _logger.LogDebug("Opening write stream for: {Path}", path);

        try
        {
            // Ensure directory exists
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            return await Task.Run(() =>
                new FileStream(
                    path,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.None,
                    DefaultBufferSize,
                    useAsync: true),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error opening write stream for: {Path}", path);
            throw;
        }
    }

    public async Task<DocumentMetadata> GetMetadataAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}", path);

        return await Task.Run(() =>
        {
            var fileInfo = new FileInfo(path);
            return new DocumentMetadata
            {
                FileName = fileInfo.Name,
                FullPath = fileInfo.FullName,
                Size = fileInfo.Length,
                CreatedDate = fileInfo.CreationTimeUtc,
                ModifiedDate = fileInfo.LastWriteTimeUtc,
                Extension = fileInfo.Extension,
                IsReadOnly = fileInfo.IsReadOnly
            };
        }, cancellationToken).ConfigureAwait(false);
    }

    public async Task SaveDocumentAsync(
        string path,
        byte[] content,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        ArgumentNullException.ThrowIfNull(content);

        _logger.LogDebug("Saving document to: {Path}", path);

        try
        {
            await using var stream = await OpenWriteStreamAsync(path, cancellationToken).ConfigureAwait(false);
            await stream.WriteAsync(content, 0, content.Length, cancellationToken).ConfigureAwait(false);
            await stream.FlushAsync(cancellationToken).ConfigureAwait(false);

            _logger.LogInformation("Document saved successfully to: {Path} ({Size} bytes)", path, content.Length);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error saving document to: {Path}", path);
            throw;
        }
    }

    public async Task<IReadOnlyList<string>> ListDocumentsAsync(
        string directory,
        string pattern,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(directory))
            throw new ArgumentNullException(nameof(directory));

        if (string.IsNullOrEmpty(pattern))
            throw new ArgumentNullException(nameof(pattern));

        if (!Directory.Exists(directory))
        {
            _logger.LogWarning("Directory not found: {Directory}", directory);
            return Array.Empty<string>();
        }

        _logger.LogDebug("Listing documents in {Directory} with pattern {Pattern}", directory, pattern);

        try
        {
            return await Task.Run(() =>
                Directory.GetFiles(directory, pattern, SearchOption.AllDirectories)
                    .OrderBy(f => f)
                    .ToList(),
                cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error listing documents in {Directory}", directory);
            throw;
        }
    }

    public async Task DeleteAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        if (!File.Exists(path))
        {
            _logger.LogWarning("File not found for deletion: {Path}", path);
            return;
        }

        _logger.LogDebug("Deleting file: {Path}", path);

        try
        {
            await Task.Run(() => File.Delete(path), cancellationToken).ConfigureAwait(false);
            _logger.LogInformation("File deleted successfully: {Path}", path);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error deleting file: {Path}", path);
            throw;
        }
    }

    public async Task<string> CreateTempDirectoryAsync(CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempPath);
            _logger.LogDebug("Created temporary directory: {Path}", tempPath);
            return tempPath;
        }, cancellationToken).ConfigureAwait(false);
    }

    public async Task<long> GetFileSizeAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}", path);

        return await Task.Run(() => new FileInfo(path).Length, cancellationToken).ConfigureAwait(false);
    }
}
