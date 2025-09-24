using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Services;

public interface IXmlParser
{
    Task<XDocument> ParseAsync(string xml, CancellationToken cancellationToken = default);
    Task<XDocument> ParseAsync(Stream stream, CancellationToken cancellationToken = default);
    Task<XDocument> LoadAsync(string filePath, CancellationToken cancellationToken = default);
}

public class XmlParser : IXmlParser
{
    private readonly ILogger<XmlParser> _logger;
    private static readonly XmlReaderSettings DefaultSettings = new()
    {
        Async = true,
        IgnoreWhitespace = false,
        IgnoreComments = true,
        DtdProcessing = DtdProcessing.Ignore,
        CheckCharacters = false
    };

    public XmlParser(ILogger<XmlParser> logger)
    {
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;
    }

    public async Task<XDocument> ParseAsync(
        string xml,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(xml))
            throw new ArgumentNullException(nameof(xml));

        try
        {
            return await Task.Run(() => XDocument.Parse(xml), cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error parsing XML string");
            throw;
        }
    }

    public async Task<XDocument> ParseAsync(
        Stream stream,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);

        try
        {
            using var reader = XmlReader.Create(stream, DefaultSettings);
            return await XDocument.LoadAsync(reader, LoadOptions.None, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error parsing XML stream");
            throw;
        }
    }

    public async Task<XDocument> LoadAsync(
        string filePath,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"XML file not found: {filePath}", filePath);

        _logger.LogDebug("Loading XML from file: {FilePath}", filePath);

        try
        {
            await using var stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                bufferSize: 4096,
                useAsync: true);

            return await ParseAsync(stream, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error loading XML from file: {FilePath}", filePath);
            throw;
        }
    }
}
