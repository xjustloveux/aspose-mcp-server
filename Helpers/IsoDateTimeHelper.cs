using System.Globalization;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Parses date and time inputs without discarding the zone the caller wrote (LOW-09).
///     <para>
///         <c>DateTime.Parse</c> with <see cref="DateTimeStyles.None" /> converts an instant
///         written as <c>2026-09-04T10:00:00Z</c> into the server's local time and marks it
///         <see cref="DateTimeKind.Local" />. Aspose then serialises it without the UTC marker,
///         so the file says 18:00 with no zone on a machine in UTC+8 and every reader elsewhere
///         sees a different instant. The intended instant is lost in the file, not just in memory.
///     </para>
///     <para>
///         Here an explicit zone (a trailing <c>Z</c> or a numeric offset) always yields
///         <see cref="DateTimeKind.Utc" />, and a value written without a zone stays
///         <see cref="DateTimeKind.Unspecified" /> so a naive local time is not silently moved.
///     </para>
/// </summary>
public static class IsoDateTimeHelper
{
    /// <summary>Parses a date/time, preserving an explicit zone as UTC.</summary>
    /// <param name="value">The date/time text, ISO 8601 preferred.</param>
    /// <param name="paramName">Parameter name reported in the exception.</param>
    /// <returns>The parsed value: UTC when the input carried a zone, otherwise unspecified.</returns>
    /// <exception cref="ArgumentException">Thrown when the value cannot be parsed.</exception>
    public static DateTime Parse(string value, string paramName)
    {
        if (!TryParse(value, out var result))
            throw new ArgumentException($"{paramName} is not a valid date/time value.", paramName);

        return result;
    }

    /// <summary>Parses a date/time, preserving an explicit zone as UTC.</summary>
    /// <param name="value">The date/time text, ISO 8601 preferred.</param>
    /// <param name="result">The parsed value when parsing succeeds.</param>
    /// <returns><c>true</c> when the value was parsed; otherwise <c>false</c>.</returns>
    public static bool TryParse(string? value, out DateTime result)
    {
        result = default;
        if (string.IsNullOrWhiteSpace(value)) return false;

        if (!DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind,
                out var parsed))
            return false;

        // RoundtripKind keeps a trailing Z as Utc and a zone-less value as Unspecified. A numeric
        // offset arrives as Local, already shifted to this machine's zone, so converting it back
        // to UTC records the instant the caller actually named.
        result = parsed.Kind == DateTimeKind.Local ? parsed.ToUniversalTime() : parsed;
        return true;
    }
}
