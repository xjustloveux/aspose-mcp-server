using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Helper for value type conversion and JSON parsing for nested structures.
///     Used when MCP SDK strongly-typed parameters don't cover complex nested JSON objects.
/// </summary>
public static class ValueHelper
{
    /// <summary>
    ///     Date shapes accepted by <see cref="LooksLikeDate" />: an ISO date, optionally with a
    ///     time part, or a day-month-year / month-day-year triple carrying a four-digit year.
    /// </summary>
    private static readonly Regex DateShapes = new(
        @"^\d{4}-\d{1,2}-\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?$"
        + @"|^\d{1,2}[/-]\d{1,2}[/-]\d{4}([ T]\d{1,2}:\d{2}(:\d{2})?)?$",
        RegexOptions.Compiled, TimeSpan.FromSeconds(2));

    /// <summary>
    ///     Parses a string value to appropriate type (number, boolean, date, or string).
    ///     Useful for Excel cell values and building typed collections.
    /// </summary>
    /// <param name="value">String value to parse.</param>
    /// <param name="asText">
    ///     When <c>true</c> the string is stored exactly as written, with no type detection. Use it
    ///     for identifiers, part numbers and anything else that merely looks numeric.
    /// </param>
    /// <returns>Parsed value as double, bool, DateTime, or original string.</returns>
    public static object ParseValue(string value, bool asText = false)
    {
        if (asText) return value;

        // NumberStyles.Any accepted accounting negatives, thousands separators and currency
        // symbols, so "(100)" became -100 and "1,234" became 1234 with no way for the caller to
        // say they meant the text. Only a plain decimal number with an optional sign and exponent
        // is treated as numeric now.
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var numValue))
            return numValue;
        if (bool.TryParse(value, out var boolValue))
            return boolValue;

        // Date detection is limited to unambiguous ISO-like forms. "3-4" used to parse as a date
        // in the current year, turning a written label into a timestamp.
        if (LooksLikeDate(value) && IsoDateTimeHelper.TryParse(value, out var dateValue))
            return dateValue;

        return value;
    }

    /// <summary>
    ///     Whether a string is shaped like a date the parser may safely interpret.
    ///     Requires a four-digit year and a full year-month-day or day-month-year triple, so a
    ///     two-part fragment such as "3-4" stays text.
    /// </summary>
    /// <param name="value">The candidate string.</param>
    /// <returns><c>true</c> when the string is an unambiguous date.</returns>
    private static bool LooksLikeDate(string value)
    {
        var text = value.Trim();
        if (text.Length < 8) return false;

        return DateShapes.IsMatch(text);
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
