using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;

namespace SharedXmlToJsonl.Providers;

public interface IOpenXmlProvider
{
    Task<Package> OpenPackageAsync(
        string path,
        OpenSettings settings,
        CancellationToken cancellationToken = default);

    Task<PackagePart?> GetPartAsync(
        Package package,
        string partUri,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<PackageRelationship>> GetRelationshipsAsync(
        PackagePart part,
        CancellationToken cancellationToken = default);

    Task<Stream> GetPartStreamAsync(
        PackagePart part,
        CancellationToken cancellationToken = default);

    Task<string> GetPartContentAsync(
        PackagePart part,
        CancellationToken cancellationToken = default);
}

public class OpenSettings
{
    public FileMode FileMode { get; set; } = FileMode.Open;
    public FileAccess FileAccess { get; set; } = FileAccess.Read;
    public FileShare FileShare { get; set; } = FileShare.Read;
    public bool ValidateOnOpen { get; set; } = true;
    public int BufferSize { get; set; } = 4096;
}