using System.Globalization;
using System.Text.Json.Nodes;

namespace AsposeMcpServer.Core.Helpers;

/// <summary>
///     Helper for value type conversion and JSON parsing for nested structures.
///     Used when MCP SDK strongly-typed parameters don't cover complex nested JSON objects.
/// </summary>
public static class ValueHelper
{
    /// <summary>
    ///     Parses a string value to appropriate type (number, boolean, date, or string).
    ///     Useful for Excel cell values and building typed collections.
    /// </summary>
    /// <param name="value">String value to parse.</param>
    /// <returns>Parsed value as double, bool, DateTime, or original string.</returns>
    public static object ParseValue(string value)
    {
        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var numValue))
            return numValue;
        if (bool.TryParse(value, out var boolValue))
            return boolValue;
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateValue))
            return dateValue;
        return value;
    }

    /// <summary>
    ///     Gets an optional JSON array from a nested JSON object.
    /// </summary>
    /// <param name="obj">JSON object to extract from.</param>
    /// <param name="key">Property key.</param>
    /// <returns>JsonArray or null if missing.</returns>
    public static JsonArray? GetArray(JsonObject? obj, string key)
    {
        if (obj == null) return null;
        var node = obj[key];
        return node as JsonArray;
    }

    /// <summary>
    ///     Gets a string from a nested JSON object with optional default value.
    /// </summary>
    /// <param name="obj">JSON object to extract from.</param>
    /// <param name="key">Property key.</param>
    /// <param name="defaultValue">Default value if missing.</param>
    /// <returns>String value or default.</returns>
    public static string GetString(JsonObject? obj, string key, string defaultValue = "")
    {
        if (obj == null) return defaultValue;
        var node = obj[key];
        return node?.GetValue<string>() ?? defaultValue;
    }
}