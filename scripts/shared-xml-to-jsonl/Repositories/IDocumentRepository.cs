using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SharedXmlToJsonl.Repositories;

public interface IDocumentRepository
{
    Task<bool> ExistsAsync(string path, CancellationToken cancellationToken = default);

    Task<Stream> OpenReadStreamAsync(string path, CancellationToken cancellationToken = default);

    Task<Stream> OpenWriteStreamAsync(string path, CancellationToken cancellationToken = default);

    Task<DocumentMetadata> GetMetadataAsync(string path, CancellationToken cancellationToken = default);

    Task SaveDocumentAsync(string path, byte[] content, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<string>> ListDocumentsAsync(
        string directory,
        string pattern,
        CancellationToken cancellationToken = default);

    Task DeleteAsync(string path, CancellationToken cancellationToken = default);

    Task<string> CreateTempDirectoryAsync(CancellationToken cancellationToken = default);

    Task<long> GetFileSizeAsync(string path, CancellationToken cancellationToken = default);
}

public class DocumentMetadata
{
    public string FileName { get; set; } = string.Empty;
    public string FullPath { get; set; } = string.Empty;
    public long Size { get; set; }
    public DateTime CreatedDate { get; set; }
    public DateTime ModifiedDate { get; set; }
    public string Extension { get; set; } = string.Empty;
    public bool IsReadOnly { get; set; }
}