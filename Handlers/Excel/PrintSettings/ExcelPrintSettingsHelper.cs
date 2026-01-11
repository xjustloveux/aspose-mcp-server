using Aspose.Cells;

namespace AsposeMcpServer.Handlers.Excel.PrintSettings;

/// <summary>
///     Helper class for Excel print settings operations.
///     Contains shared functionality used by multiple print settings handlers.
/// </summary>
public static class ExcelPrintSettingsHelper
{
    /// <summary>
    ///     Supported paper size mappings.
    /// </summary>
    public static readonly Dictionary<string, PaperSizeType> PaperSizeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["A3"] = PaperSizeType.PaperA3,
        ["A4"] = PaperSizeType.PaperA4,
        ["A5"] = PaperSizeType.PaperA5,
        ["B4"] = PaperSizeType.PaperB4,
        ["B5"] = PaperSizeType.PaperB5,
        ["Letter"] = PaperSizeType.PaperLetter,
        ["Legal"] = PaperSizeType.PaperLegal,
        ["Tabloid"] = PaperSizeType.PaperTabloid,
        ["Executive"] = PaperSizeType.PaperExecutive
    };

    /// <summary>
    ///     Applies page setup options to the PageSetup object.
    /// </summary>
    /// <param name="pageSetup">The PageSetup object to modify.</param>
    /// <param name="orientation">The page orientation ('Portrait' or 'Landscape').</param>
    /// <param name="paperSize">The paper size (e.g., 'A4', 'Letter').</param>
    /// <param name="leftMargin">The left margin in inches.</param>
    /// <param name="rightMargin">The right margin in inches.</param>
    /// <param name="topMargin">The top margin in inches.</param>
    /// <param name="bottomMargin">The bottom margin in inches.</param>
    /// <param name="header">The header text for center section.</param>
    /// <param name="footer">The footer text for center section.</param>
    /// <param name="fitToPage">Whether to enable fit to page mode.</param>
    /// <param name="fitToPagesWide">The number of pages wide to fit content.</param>
    /// <param name="fitToPagesTall">The number of pages tall to fit content.</param>
    /// <returns>A list of change descriptions indicating what settings were modified.</returns>
    /// <exception cref="ArgumentException">Thrown when an invalid paper size is specified.</exception>
    public static List<string> ApplyPageSetup(PageSetup pageSetup, string? orientation, string? paperSize,
        double? leftMargin, double? rightMargin, double? topMargin, double? bottomMargin,
        string? header, string? footer, bool? fitToPage, int? fitToPagesWide, int? fitToPagesTall)
    {
        List<string> changes = [];

        if (!string.IsNullOrEmpty(orientation))
        {
            pageSetup.Orientation = string.Equals(orientation, "Landscape", StringComparison.OrdinalIgnoreCase)
                ? PageOrientationType.Landscape
                : PageOrientationType.Portrait;
            changes.Add($"orientation={orientation}");
        }

        if (!string.IsNullOrEmpty(paperSize))
        {
            if (PaperSizeMap.TryGetValue(paperSize, out var size))
            {
                pageSetup.PaperSize = size;
                changes.Add($"paperSize={paperSize}");
            }
            else
            {
                throw new ArgumentException(
                    $"Invalid paper size: '{paperSize}'. Supported values: {string.Join(", ", PaperSizeMap.Keys)}");
            }
        }

        if (leftMargin.HasValue)
        {
            pageSetup.LeftMargin = leftMargin.Value;
            changes.Add($"leftMargin={leftMargin.Value}");
        }

        if (rightMargin.HasValue)
        {
            pageSetup.RightMargin = rightMargin.Value;
            changes.Add($"rightMargin={rightMargin.Value}");
        }

        if (topMargin.HasValue)
        {
            pageSetup.TopMargin = topMargin.Value;
            changes.Add($"topMargin={topMargin.Value}");
        }

        if (bottomMargin.HasValue)
        {
            pageSetup.BottomMargin = bottomMargin.Value;
            changes.Add($"bottomMargin={bottomMargin.Value}");
        }

        if (!string.IsNullOrEmpty(header))
        {
            pageSetup.SetHeader(1, header);
            changes.Add("header");
        }

        if (!string.IsNullOrEmpty(footer))
        {
            pageSetup.SetFooter(1, footer);
            changes.Add("footer");
        }

        if (fitToPage == true)
        {
            pageSetup.FitToPagesWide = fitToPagesWide ?? 1;
            pageSetup.FitToPagesTall = fitToPagesTall ?? 1;
            changes.Add($"fitToPage(wide={pageSetup.FitToPagesWide}, tall={pageSetup.FitToPagesTall})");
        }

        return changes;
    }
}
