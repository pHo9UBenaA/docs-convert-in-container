using System;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using SharedXmlToJsonl.Models;

namespace SharedXmlToJsonl;

/// <summary>
/// Custom JSON serializer context with proper encoding settings
/// </summary>
public static class JsonSerializationContext
{
    /// <summary>
    /// Thread-safe lazy initialization of the JSON serializer context
    /// </summary>
    private static readonly Lazy<ElementJsonSerializerContext> LazyContext = new(() =>
    {
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
            WriteIndented = false,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };
        return new ElementJsonSerializerContext(options);
    });

    /// <summary>
    /// Gets the default JSON serializer context with Japanese text support
    /// </summary>
    public static ElementJsonSerializerContext Default => LazyContext.Value;
}
