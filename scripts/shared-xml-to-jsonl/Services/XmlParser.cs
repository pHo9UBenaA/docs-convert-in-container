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

public partial class XmlParser : IXmlParser
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
            LogErrorParsingXmlString(_logger, ex);
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
            LogErrorParsingXmlStream(_logger, ex);
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

        LogLoadingXmlFromFile(_logger, filePath);

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
            LogErrorLoadingXmlFromFile(_logger, ex, filePath);
            throw;
        }
    }

    [LoggerMessage(
        EventId = 5001,
        Level = LogLevel.Error,
        Message = "Error parsing XML string")]
    private static partial void LogErrorParsingXmlString(
        ILogger logger, Exception ex);

    [LoggerMessage(
        EventId = 5002,
        Level = LogLevel.Error,
        Message = "Error parsing XML stream")]
    private static partial void LogErrorParsingXmlStream(
        ILogger logger, Exception ex);

    [LoggerMessage(
        EventId = 5003,
        Level = LogLevel.Debug,
        Message = "Loading XML from file: {filePath}")]
    private static partial void LogLoadingXmlFromFile(
        ILogger logger, string filePath);

    [LoggerMessage(
        EventId = 5004,
        Level = LogLevel.Error,
        Message = "Error loading XML from file: {filePath}")]
    private static partial void LogErrorLoadingXmlFromFile(
        ILogger logger, Exception ex, string filePath);
}
