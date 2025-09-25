using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using SharedXmlToJsonl.Interfaces;
using SharedXmlToJsonl.Models;

namespace SharedXmlToJsonl.Services;

public partial class JsonWriter : IJsonWriter
{
    private readonly ILogger<JsonWriter> _logger;
    private static readonly JsonSerializerOptions JsonSerializerOptions = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        WriteIndented = false,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
        TypeInfoResolver = SharedJsonSerializerContext.Default
    };

    public JsonWriter(ILogger<JsonWriter> logger)
    {
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;
    }

    public async Task WriteJsonLineAsync<T>(
        StreamWriter writer,
        T obj,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(writer);

        ArgumentNullException.ThrowIfNull(obj);

        try
        {
            string json;
            if (obj is DocumentEntry documentEntry)
            {
                json = JsonSerializer.Serialize(documentEntry, SharedJsonSerializerContext.Default.DocumentEntry);
            }
            else if (obj is ProcessingResult processingResult)
            {
                json = JsonSerializer.Serialize(processingResult, SharedJsonSerializerContext.Default.ProcessingResult);
            }
            else
            {
                json = JsonSerializer.Serialize(obj, JsonSerializerOptions);
            }
            await writer.WriteLineAsync(json.AsMemory(), cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogErrorSerializingObjectToJson(_logger, ex);
            throw;
        }
    }

    public async Task WriteJsonLinesAsync<T>(
        string filePath,
        IEnumerable<T> objects,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));

        ArgumentNullException.ThrowIfNull(objects);

        LogWritingJsonLinesToFile(_logger, filePath);

        try
        {
            // Ensure directory exists
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            await using var fileStream = new FileStream(
                filePath,
                FileMode.Create,
                FileAccess.Write,
                FileShare.None,
                bufferSize: 4096,
                useAsync: true);

            await using var writer = new StreamWriter(fileStream, CommonBase.Utf8NoBom);

            var count = 0;
            foreach (var obj in objects)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await WriteJsonLineAsync(writer, obj, cancellationToken).ConfigureAwait(false);
                count++;
            }

            await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
            LogSuccessfullyWroteJsonLines(_logger, count, filePath);
        }
        catch (OperationCanceledException)
        {
            LogJsonWritingCancelled(_logger, filePath);
            throw;
        }
        catch (Exception ex)
        {
            LogErrorWritingJsonLinesToFile(_logger, ex, filePath);
            throw;
        }
    }

    [LoggerMessage(
        EventId = 8001,
        Level = LogLevel.Error,
        Message = "Error serializing object to JSON")]
    private static partial void LogErrorSerializingObjectToJson(
        ILogger logger, Exception ex);

    [LoggerMessage(
        EventId = 8002,
        Level = LogLevel.Debug,
        Message = "Writing JSON lines to {filePath}")]
    private static partial void LogWritingJsonLinesToFile(
        ILogger logger, string filePath);

    [LoggerMessage(
        EventId = 8003,
        Level = LogLevel.Information,
        Message = "Successfully wrote {count} JSON lines to {filePath}")]
    private static partial void LogSuccessfullyWroteJsonLines(
        ILogger logger, int count, string filePath);

    [LoggerMessage(
        EventId = 8004,
        Level = LogLevel.Warning,
        Message = "JSON writing was cancelled for {filePath}")]
    private static partial void LogJsonWritingCancelled(
        ILogger logger, string filePath);

    [LoggerMessage(
        EventId = 8005,
        Level = LogLevel.Error,
        Message = "Error writing JSON lines to {filePath}")]
    private static partial void LogErrorWritingJsonLinesToFile(
        ILogger logger, Exception ex, string filePath);
}
