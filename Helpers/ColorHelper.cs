using System.Drawing;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Helper class for color parsing and conversion
/// </summary>
public static class ColorHelper
{
    /// <summary>
    ///     Parses a color string to System.Drawing.Color
    ///     Supports formats:
    ///     - Hex: #RRGGBB, RRGGBB, #AARRGGBB, AARRGGBB
    ///     - RGB comma-separated: "255,0,0" or "255, 0, 0"
    ///     - Named colors: "Red", "Blue", etc.
    ///     Returns Color.Black if parsing fails or string is null/empty.
    /// </summary>
    /// <param name="colorString">Color string in various formats</param>
    /// <returns>System.Drawing.Color object, or Color.Black if parsing fails</returns>
    public static Color ParseColor(string? colorString)
    {
        if (string.IsNullOrWhiteSpace(colorString))
            return Color.Black;

        var result = ParseColorInternal(colorString);
        return result ?? Color.Black;
    }

    /// <summary>
    ///     Parses a color string with a custom default color
    /// </summary>
    /// <param name="colorString">Color string in various formats</param>
    /// <param name="defaultColor">Default color to return if parsing fails or string is null/empty</param>
    /// <returns>Parsed color or default color</returns>
    public static Color ParseColor(string? colorString, Color defaultColor)
    {
        if (string.IsNullOrWhiteSpace(colorString))
            return defaultColor;

        var result = ParseColorInternal(colorString);
        return result ?? defaultColor;
    }

    /// <summary>
    ///     Parses a color string and throws ArgumentException if parsing fails
    ///     This overload uses a marker type to indicate that exceptions should be thrown
    /// </summary>
    /// <param name="colorString">Color string in various formats</param>
    /// <param name="throwOnError">Marker parameter - when true, throws ArgumentException on parse failure</param>
    /// <returns>Parsed color</returns>
    /// <exception cref="ArgumentException">Thrown when parsing fails or string is null/empty</exception>
    public static Color ParseColor(string? colorString, bool throwOnError)
    {
        if (throwOnError)
        {
            if (string.IsNullOrWhiteSpace(colorString))
                throw new ArgumentException("Color string cannot be null or empty", nameof(colorString));

            var result = ParseColorInternal(colorString);
            if (result.HasValue)
                return result.Value;

            throw new ArgumentException(
                $"Unable to parse color '{colorString}'. Please use a valid color format (e.g., #FF0000, 255,0,0, or red)",
                nameof(colorString));
        }

        return ParseColor(colorString);
    }

    /// <summary>
    ///     Tries to parse a color string to System.Drawing.Color
    ///     Returns true if parsing succeeds, false otherwise
    /// </summary>
    /// <param name="colorString">Color string in various formats</param>
    /// <param name="color">Parsed color if successful, otherwise Color.Black</param>
    /// <returns>True if parsing succeeded, false otherwise</returns>
    public static bool TryParseColor(string? colorString, out Color color)
    {
        if (string.IsNullOrWhiteSpace(colorString))
        {
            color = Color.Black;
            return false;
        }

        var result = ParseColorInternal(colorString);
        if (result.HasValue)
        {
            color = result.Value;
            return true;
        }

        color = Color.Black;
        return false;
    }

    /// <summary>
    ///     Internal method that performs the actual color parsing.
    /// </summary>
    /// <param name="colorString">The color string to parse.</param>
    /// <returns>The parsed color, or null if parsing fails.</returns>
    private static Color? ParseColorInternal(string colorString)
    {
        colorString = colorString.Trim();

        if (colorString.Contains(','))
        {
            var parts = colorString.Split(',');
            if (parts.Length == 3 &&
                int.TryParse(parts[0].Trim(), out var r) &&
                int.TryParse(parts[1].Trim(), out var g) &&
                int.TryParse(parts[2].Trim(), out var b))
            {
                r = Math.Max(0, Math.Min(255, r));
                g = Math.Max(0, Math.Min(255, g));
                b = Math.Max(0, Math.Min(255, b));
                return Color.FromArgb(r, g, b);
            }
        }

        try
        {
            var hexColor = colorString.TrimStart('#');

            if (hexColor.Length == 8)
            {
                var a = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                var r = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                var g = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                var b = Convert.ToInt32(hexColor.Substring(6, 2), 16);
                return Color.FromArgb(a, r, g, b);
            }

            if (hexColor.Length == 6)
            {
                var r = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                var g = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                var b = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                return Color.FromArgb(r, g, b);
            }
        }
        catch
        {
            // Ignore hex parsing errors, try named color next
        }

        try
        {
            var namedColor = Color.FromName(colorString);
            if (namedColor.IsKnownColor || namedColor.A != 0)
                return namedColor;
        }
        catch
        {
            // Ignore named color parsing errors
        }

        return null;
    }

    /// <summary>
    ///     Converts System.Drawing.Color to Aspose.Pdf.Color
    /// </summary>
    /// <param name="color">System.Drawing.Color object to convert</param>
    /// <returns>Aspose.Pdf.Color object with RGB values normalized to 0-1 range</returns>
    public static Aspose.Pdf.Color ToPdfColor(Color color)
    {
        return Aspose.Pdf.Color.FromRgb(
            color.R / 255f,
            color.G / 255f,
            color.B / 255f
        );
    }
}
