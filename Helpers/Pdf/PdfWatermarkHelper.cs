using Aspose.Pdf;

namespace AsposeMcpServer.Helpers.Pdf;

/// <summary>
///     Helper methods for PDF watermark operations.
/// </summary>
public static class PdfWatermarkHelper
{
    /// <summary>
    ///     Parses a color name or hex code into a PDF Color object.
    /// </summary>
    /// <param name="colorName">The color name or hex code (e.g., "Red" or "#FF0000").</param>
    /// <returns>The parsed Color object, or Gray if parsing fails.</returns>
    public static Color ParseColor(string colorName)
    {
        if (string.IsNullOrEmpty(colorName))
            return Color.Gray;

        if (colorName.StartsWith('#') && (colorName.Length == 7 || colorName.Length == 9))
            try
            {
                var hex = colorName.TrimStart('#');
                var r = Convert.ToByte(hex[..2], 16);
                var g = Convert.ToByte(hex.Substring(2, 2), 16);
                var b = Convert.ToByte(hex.Substring(4, 2), 16);
                return Color.FromRgb(r / 255.0, g / 255.0, b / 255.0);
            }
            catch
            {
                return Color.Gray;
            }

        return colorName.ToLower() switch
        {
            "red" => Color.Red,
            "blue" => Color.Blue,
            "green" => Color.Green,
            "black" => Color.Black,
            "white" => Color.White,
            "yellow" => Color.Yellow,
            "orange" => Color.Orange,
            "purple" => Color.Purple,
            "pink" => Color.Pink,
            "cyan" => Color.Cyan,
            "magenta" => Color.Magenta,
            "lightgray" => Color.LightGray,
            "darkgray" => Color.DarkGray,
            _ => Color.Gray
        };
    }

    /// <summary>
    ///     Parses a page range string into a list of page indices.
    /// </summary>
    /// <param name="pageRange">The page range string (e.g., "1,3,5-10").</param>
    /// <param name="totalPages">The total number of pages in the document.</param>
    /// <returns>A list of 1-based page indices.</returns>
    /// <exception cref="ArgumentException">Thrown when the page range format is invalid or out of bounds.</exception>
    public static List<int> ParsePageRange(string? pageRange, int totalPages)
    {
        if (string.IsNullOrEmpty(pageRange))
            return Enumerable.Range(1, totalPages).ToList();

        var result = new HashSet<int>();
        var parts = pageRange.Split(',', StringSplitOptions.RemoveEmptyEntries);

        foreach (var part in parts)
        {
            var trimmed = part.Trim();
            if (trimmed.Contains('-'))
                ParseRangePart(trimmed, totalPages, result);
            else
                ParseSinglePage(trimmed, totalPages, result);
        }

        return result.OrderBy(x => x).ToList();
    }

    /// <summary>
    ///     Parses a range part (e.g., "5-10") and adds page indices to the result.
    /// </summary>
    /// <param name="trimmed">The range string to parse.</param>
    /// <param name="totalPages">The total number of pages in the document.</param>
    /// <param name="result">The set to add page indices to.</param>
    private static void ParseRangePart(string trimmed, int totalPages, HashSet<int> result)
    {
        var rangeParts = trimmed.Split('-');
        if (rangeParts.Length != 2 ||
            !int.TryParse(rangeParts[0].Trim(), out var start) ||
            !int.TryParse(rangeParts[1].Trim(), out var end))
            throw new ArgumentException(
                $"Invalid page range format: '{trimmed}'. Expected format: 'start-end' (e.g., '5-10')");

        if (start < 1 || end > totalPages || start > end)
            throw new ArgumentException(
                $"Page range '{trimmed}' is out of bounds. Document has {totalPages} page(s)");

        for (var i = start; i <= end; i++)
            result.Add(i);
    }

    /// <summary>
    ///     Parses a single page number and adds it to the result.
    /// </summary>
    /// <param name="trimmed">The page number string to parse.</param>
    /// <param name="totalPages">The total number of pages in the document.</param>
    /// <param name="result">The set to add the page index to.</param>
    private static void ParseSinglePage(string trimmed, int totalPages, HashSet<int> result)
    {
        if (!int.TryParse(trimmed, out var pageNum))
            throw new ArgumentException($"Invalid page number: '{trimmed}'");

        if (pageNum < 1 || pageNum > totalPages)
            throw new ArgumentException(
                $"Page number {pageNum} is out of bounds. Document has {totalPages} page(s)");

        result.Add(pageNum);
    }
}
