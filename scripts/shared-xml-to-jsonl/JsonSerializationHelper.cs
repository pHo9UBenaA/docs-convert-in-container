using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;

namespace SharedXmlToJsonl;

/// <summary>
/// Provides common JSON serialization utilities for JSONL output
/// </summary>
public static class JsonSerializationHelper
{
    /// <summary>
    /// Gets the standard JSON serializer options for JSONL output
    /// </summary>
    public static JsonSerializerOptions CreateStandardOptions()
    {
        return new JsonSerializerOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            WriteIndented = false,
            PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower
        };
    }

    /// <summary>
    /// Base class for source generation contexts
    /// Derived classes should be partial and have JsonSerializable attributes
    /// </summary>
    public abstract class BaseSourceGenerationContext : JsonSerializerContext
    {
        private static readonly JsonSerializerOptions StandardOptions = CreateStandardOptions();

        protected BaseSourceGenerationContext() : base(StandardOptions)
        {
        }

        protected BaseSourceGenerationContext(JsonSerializerOptions options) : base(options)
        {
        }
    }

    /// <summary>
    /// Writes an object as a single JSONL line to a StreamWriter
    /// </summary>
    public static void WriteAsJsonLine<T>(StreamWriter writer, T obj, JsonSerializerContext context)
    {
        var json = JsonSerializer.Serialize(obj, typeof(T), context);
        writer.WriteLine(json);
    }

    /// <summary>
    /// Writes an object as a single JSONL line with a specific type info
    /// </summary>
    public static void WriteAsJsonLine<T>(StreamWriter writer, T obj, JsonTypeInfo<T> typeInfo)
    {
        var json = JsonSerializer.Serialize(obj, typeInfo);
        writer.WriteLine(json);
    }

    /// <summary>
    /// Serializes an object to a JSON string using the provided context
    /// </summary>
    public static string SerializeToJson<T>(T obj, JsonSerializerContext context)
    {
        return JsonSerializer.Serialize(obj, typeof(T), context);
    }

    /// <summary>
    /// Serializes an object to a JSON string with a specific type info
    /// </summary>
    public static string SerializeToJson<T>(T obj, JsonTypeInfo<T> typeInfo)
    {
        return JsonSerializer.Serialize(obj, typeInfo);
    }
}
