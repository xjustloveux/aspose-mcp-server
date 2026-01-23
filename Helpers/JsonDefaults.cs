using System.Text.Json;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Provides shared JSON serialization options for consistent serialization across the application.
/// </summary>
public static class JsonDefaults
{
    /// <summary>
    ///     Default JSON serializer options with camelCase naming policy.
    /// </summary>
    public static JsonSerializerOptions CamelCase { get; } = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    /// <summary>
    ///     JSON serializer options with camelCase naming policy and indented formatting.
    /// </summary>
    public static JsonSerializerOptions CamelCaseIndented { get; } = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    /// <summary>
    ///     JSON serializer options with indented formatting only.
    /// </summary>
    public static JsonSerializerOptions Indented { get; } = new()
    {
        WriteIndented = true
    };

    /// <summary>
    ///     JSON serializer options with case-insensitive property name matching for deserialization.
    /// </summary>
    public static JsonSerializerOptions CaseInsensitive { get; } = new()
    {
        PropertyNameCaseInsensitive = true
    };
}
