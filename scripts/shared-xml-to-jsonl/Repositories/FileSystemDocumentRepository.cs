using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Repositories;

public partial class FileSystemDocumentRepository : IDocumentRepository
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

        LogOpeningReadStream(_logger, path);

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
            LogErrorOpeningReadStream(_logger, ex, path);
            throw;
        }
    }

    public async Task<Stream> OpenWriteStreamAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        LogOpeningWriteStream(_logger, path);

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
            LogErrorOpeningWriteStream(_logger, ex, path);
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

        LogSavingDocument(_logger, path);

        try
        {
            await using var stream = await OpenWriteStreamAsync(path, cancellationToken).ConfigureAwait(false);
            await stream.WriteAsync(content, 0, content.Length, cancellationToken).ConfigureAwait(false);
            await stream.FlushAsync(cancellationToken).ConfigureAwait(false);

            LogDocumentSavedSuccessfully(_logger, path, content.Length);
        }
        catch (Exception ex)
        {
            LogErrorSavingDocument(_logger, ex, path);
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
            LogDirectoryNotFound(_logger, directory);
            return Array.Empty<string>();
        }

        LogListingDocuments(_logger, directory, pattern);

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
            LogErrorListingDocuments(_logger, ex, directory);
            throw;
        }
    }

    public async Task DeleteAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentNullException(nameof(path));

        if (!File.Exists(path))
        {
            LogFileNotFoundForDeletion(_logger, path);
            return;
        }

        LogDeletingFile(_logger, path);

        try
        {
            await Task.Run(() => File.Delete(path), cancellationToken).ConfigureAwait(false);
            LogFileDeletedSuccessfully(_logger, path);
        }
        catch (Exception ex)
        {
            LogErrorDeletingFile(_logger, ex, path);
            throw;
        }
    }

    public async Task<string> CreateTempDirectoryAsync(CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempPath);
            LogCreatedTemporaryDirectory(_logger, tempPath);
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

    [LoggerMessage(
        EventId = 4001,
        Level = LogLevel.Debug,
        Message = "Opening read stream for: {path}")]
    private static partial void LogOpeningReadStream(
        ILogger logger, string path);

    [LoggerMessage(
        EventId = 4002,
        Level = LogLevel.Error,
        Message = "Error opening read stream for: {path}")]
    private static partial void LogErrorOpeningReadStream(
        ILogger logger, Exception ex, string path);

    [LoggerMessage(
        EventId = 4003,
        Level = LogLevel.Debug,
        Message = "Opening write stream for: {path}")]
    private static partial void LogOpeningWriteStream(
        ILogger logger, string path);

    [LoggerMessage(
        EventId = 4004,
        Level = LogLevel.Error,
        Message = "Error opening write stream for: {path}")]
    private static partial void LogErrorOpeningWriteStream(
        ILogger logger, Exception ex, string path);

    [LoggerMessage(
        EventId = 4005,
        Level = LogLevel.Debug,
        Message = "Saving document to: {path}")]
    private static partial void LogSavingDocument(
        ILogger logger, string path);

    [LoggerMessage(
        EventId = 4006,
        Level = LogLevel.Information,
        Message = "Document saved successfully to: {path} ({size} bytes)")]
    private static partial void LogDocumentSavedSuccessfully(
        ILogger logger, string path, int size);

    [LoggerMessage(
        EventId = 4007,
        Level = LogLevel.Error,
        Message = "Error saving document to: {path}")]
    private static partial void LogErrorSavingDocument(
        ILogger logger, Exception ex, string path);

    [LoggerMessage(
        EventId = 4008,
        Level = LogLevel.Warning,
        Message = "Directory not found: {directory}")]
    private static partial void LogDirectoryNotFound(
        ILogger logger, string directory);

    [LoggerMessage(
        EventId = 4009,
        Level = LogLevel.Debug,
        Message = "Listing documents in {directory} with pattern {pattern}")]
    private static partial void LogListingDocuments(
        ILogger logger, string directory, string pattern);

    [LoggerMessage(
        EventId = 4010,
        Level = LogLevel.Error,
        Message = "Error listing documents in {directory}")]
    private static partial void LogErrorListingDocuments(
        ILogger logger, Exception ex, string directory);

    [LoggerMessage(
        EventId = 4011,
        Level = LogLevel.Warning,
        Message = "File not found for deletion: {path}")]
    private static partial void LogFileNotFoundForDeletion(
        ILogger logger, string path);

    [LoggerMessage(
        EventId = 4012,
        Level = LogLevel.Debug,
        Message = "Deleting file: {path}")]
    private static partial void LogDeletingFile(
        ILogger logger, string path);

    [LoggerMessage(
        EventId = 4013,
        Level = LogLevel.Information,
        Message = "File deleted successfully: {path}")]
    private static partial void LogFileDeletedSuccessfully(
        ILogger logger, string path);

    [LoggerMessage(
        EventId = 4014,
        Level = LogLevel.Error,
        Message = "Error deleting file: {path}")]
    private static partial void LogErrorDeletingFile(
        ILogger logger, Exception ex, string path);

    [LoggerMessage(
        EventId = 4015,
        Level = LogLevel.Debug,
        Message = "Created temporary directory: {path}")]
    private static partial void LogCreatedTemporaryDirectory(
        ILogger logger, string path);
}
