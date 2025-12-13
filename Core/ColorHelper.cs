namespace AsposeMcpServer.Core;

/// <summary>
/// Helper class for color parsing and conversion
/// </summary>
public static class ColorHelper
{
    /// <summary>
    /// Parses a hex color string to System.Drawing.Color
    /// Supports formats: #RRGGBB, RRGGBB, #AARRGGBB, AARRGGBB, and named colors
    /// </summary>
    /// <param name="hexColor">Hex color string (with or without # prefix) or named color</param>
    /// <returns>System.Drawing.Color object</returns>
    public static System.Drawing.Color ParseColor(string hexColor)
    {
        if (string.IsNullOrEmpty(hexColor))
            return System.Drawing.Color.Black;

        try
        {
            hexColor = hexColor.TrimStart('#');

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
            else
            {
                // Try named color
                return System.Drawing.Color.FromName(hexColor);
            }
        }
        catch
        {
            // If parsing fails, try named color, otherwise return black
            try
            {
                return System.Drawing.Color.FromName(hexColor);
            }
            catch
            {
                return System.Drawing.Color.Black;
            }
        }
    }
}

