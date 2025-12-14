namespace AsposeMcpServer.Core;

/// <summary>
/// Helper class for color parsing and conversion
/// </summary>
public static class ColorHelper
{
    /// <summary>
    /// Parses a color string to System.Drawing.Color
    /// Supports formats:
    /// - Hex: #RRGGBB, RRGGBB, #AARRGGBB, AARRGGBB
    /// - RGB comma-separated: "255,0,0" or "255, 0, 0"
    /// - Named colors: "Red", "Blue", etc.
    /// </summary>
    /// <param name="colorString">Color string in various formats</param>
    /// <returns>System.Drawing.Color object</returns>
    public static System.Drawing.Color ParseColor(string colorString)
    {
        if (string.IsNullOrWhiteSpace(colorString))
            return System.Drawing.Color.Black;

        colorString = colorString.Trim();

        // Try RGB comma-separated format first (e.g., "255,0,0" or "255, 0, 0")
        if (colorString.Contains(','))
        {
            var parts = colorString.Split(',');
            if (parts.Length == 3)
            {
                if (int.TryParse(parts[0].Trim(), out int r) &&
                    int.TryParse(parts[1].Trim(), out int g) &&
                    int.TryParse(parts[2].Trim(), out int b))
                {
                    // Clamp values to valid range
                    r = Math.Max(0, Math.Min(255, r));
                    g = Math.Max(0, Math.Min(255, g));
                    b = Math.Max(0, Math.Min(255, b));
                    return System.Drawing.Color.FromArgb(r, g, b);
                }
            }
        }

        // Try hex format
        try
        {
            var hexColor = colorString.TrimStart('#');

            if (hexColor.Length == 8)
            {
                // ARGB format
                int a = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                int r = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                int g = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                int b = Convert.ToInt32(hexColor.Substring(6, 2), 16);
                return System.Drawing.Color.FromArgb(a, r, g, b);
            }
            else if (hexColor.Length == 6)
            {
                // RGB format
                int r = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                int g = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                int b = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                return System.Drawing.Color.FromArgb(r, g, b);
            }
        }
        catch
        {
            // Continue to try named color
        }

        // Try named color
        try
        {
            var namedColor = System.Drawing.Color.FromName(colorString);
            if (namedColor.IsKnownColor || namedColor.A != 0)
            {
                return namedColor;
            }
        }
        catch
        {
            // Fall through to default
        }

        // If all parsing fails, return black
        return System.Drawing.Color.Black;
    }

    /// <summary>
    /// Converts System.Drawing.Color to Aspose.Pdf.Color
    /// </summary>
    /// <param name="color">System.Drawing.Color object</param>
    /// <returns>Aspose.Pdf.Color object</returns>
    public static Aspose.Pdf.Color ToPdfColor(System.Drawing.Color color)
    {
        return Aspose.Pdf.Color.FromRgb(
            color.R / 255f,
            color.G / 255f,
            color.B / 255f
        );
    }
}

