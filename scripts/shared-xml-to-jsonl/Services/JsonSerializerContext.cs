using System;
using System.Text.Json.Serialization;
using System.Collections.Generic;
using SharedXmlToJsonl.Models;

namespace SharedXmlToJsonl.Services;

/// <summary>
/// JSON serializer context for source generation
/// </summary>
[JsonSerializable(typeof(DocumentEntry))]
[JsonSerializable(typeof(ProcessingResult))]
[JsonSerializable(typeof(ElementBase))]
[JsonSerializable(typeof(ShapeElement))]
[JsonSerializable(typeof(TableElement))]
[JsonSerializable(typeof(ConnectorElement))]
[JsonSerializable(typeof(PictureElement))]
[JsonSerializable(typeof(ChartElement))]
[JsonSerializable(typeof(SpreadsheetTableElement))]
[JsonSerializable(typeof(Position))]
[JsonSerializable(typeof(Size))]
[JsonSerializable(typeof(List<List<string>>))]
[JsonSerializable(typeof(List<List<object>>))]
[JsonSerializable(typeof(object))]
[JsonSerializable(typeof(string))]
[JsonSerializable(typeof(double))]
[JsonSerializable(typeof(int))]
[JsonSerializable(typeof(bool))]
[JsonSerializable(typeof(DateTime))]
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.SnakeCaseLower,
    WriteIndented = false,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
internal partial class SharedJsonSerializerContext : JsonSerializerContext
{
}