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
    private static ElementJsonSerializerContext? _context;

    /// <summary>
    /// Gets the default JSON serializer context with Japanese text support
    /// </summary>
    public static ElementJsonSerializerContext Default
    {
        get
        {
            if (_context == null)
            {
                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
                    WriteIndented = false,
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                _context = new ElementJsonSerializerContext(options);
            }
            return _context;
        }
    }
}
