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

public class JsonWriter : IJsonWriter
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
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public async Task WriteJsonLineAsync<T>(
        StreamWriter writer,
        T obj,
        CancellationToken cancellationToken = default)
    {
        if (writer == null)
            throw new ArgumentNullException(nameof(writer));

        if (obj == null)
            throw new ArgumentNullException(nameof(obj));

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
            _logger.LogError(ex, "Error serializing object to JSON");
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

        if (objects == null)
            throw new ArgumentNullException(nameof(objects));

        _logger.LogDebug("Writing JSON lines to {FilePath}", filePath);

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

            await writer.FlushAsync().ConfigureAwait(false);
            _logger.LogInformation("Successfully wrote {Count} JSON lines to {FilePath}", count, filePath);
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("JSON writing was cancelled for {FilePath}", filePath);
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error writing JSON lines to {FilePath}", filePath);
            throw;
        }
    }
}