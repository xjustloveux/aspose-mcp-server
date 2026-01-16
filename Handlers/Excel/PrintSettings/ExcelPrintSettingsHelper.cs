using System.Collections.Frozen;
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
    public static FrozenDictionary<string, PaperSizeType> PaperSizeMap { get; } =
        new Dictionary<string, PaperSizeType>(StringComparer.OrdinalIgnoreCase)
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
        }.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    ///     Applies page setup options to the PageSetup object.
    /// </summary>
    /// <param name="pageSetup">The PageSetup object to modify.</param>
    /// <param name="options">The page setup options to apply.</param>
    /// <returns>A list of change descriptions indicating what settings were modified.</returns>
    /// <exception cref="ArgumentException">Thrown when an invalid paper size is specified.</exception>
    public static List<string> ApplyPageSetup(PageSetup pageSetup, PageSetupOptions options)
    {
        List<string> changes = [];

        if (!string.IsNullOrEmpty(options.Orientation))
        {
            pageSetup.Orientation = string.Equals(options.Orientation, "Landscape", StringComparison.OrdinalIgnoreCase)
                ? PageOrientationType.Landscape
                : PageOrientationType.Portrait;
            changes.Add($"orientation={options.Orientation}");
        }

        if (!string.IsNullOrEmpty(options.PaperSize))
        {
            if (PaperSizeMap.TryGetValue(options.PaperSize, out var size))
            {
                pageSetup.PaperSize = size;
                changes.Add($"paperSize={options.PaperSize}");
            }
            else
            {
                throw new ArgumentException(
                    $"Invalid paper size: '{options.PaperSize}'. Supported values: {string.Join(", ", PaperSizeMap.Keys)}");
            }
        }

        if (options.LeftMargin.HasValue)
        {
            pageSetup.LeftMargin = options.LeftMargin.Value;
            changes.Add($"leftMargin={options.LeftMargin.Value}");
        }

        if (options.RightMargin.HasValue)
        {
            pageSetup.RightMargin = options.RightMargin.Value;
            changes.Add($"rightMargin={options.RightMargin.Value}");
        }

        if (options.TopMargin.HasValue)
        {
            pageSetup.TopMargin = options.TopMargin.Value;
            changes.Add($"topMargin={options.TopMargin.Value}");
        }

        if (options.BottomMargin.HasValue)
        {
            pageSetup.BottomMargin = options.BottomMargin.Value;
            changes.Add($"bottomMargin={options.BottomMargin.Value}");
        }

        if (!string.IsNullOrEmpty(options.Header))
        {
            pageSetup.SetHeader(1, options.Header);
            changes.Add("header");
        }

        if (!string.IsNullOrEmpty(options.Footer))
        {
            pageSetup.SetFooter(1, options.Footer);
            changes.Add("footer");
        }

        if (options.FitToPage == true)
        {
            pageSetup.FitToPagesWide = options.FitToPagesWide ?? 1;
            pageSetup.FitToPagesTall = options.FitToPagesTall ?? 1;
            changes.Add($"fitToPage(wide={pageSetup.FitToPagesWide}, tall={pageSetup.FitToPagesTall})");
        }

        return changes;
    }
}

/// <summary>
///     Record to hold page setup options.
/// </summary>
/// <param name="Orientation">The page orientation ('Portrait' or 'Landscape').</param>
/// <param name="PaperSize">The paper size (e.g., 'A4', 'Letter').</param>
/// <param name="LeftMargin">The left margin in inches.</param>
/// <param name="RightMargin">The right margin in inches.</param>
/// <param name="TopMargin">The top margin in inches.</param>
/// <param name="BottomMargin">The bottom margin in inches.</param>
/// <param name="Header">The header text for center section.</param>
/// <param name="Footer">The footer text for center section.</param>
/// <param name="FitToPage">Whether to enable fit to page mode.</param>
/// <param name="FitToPagesWide">The number of pages wide to fit content.</param>
/// <param name="FitToPagesTall">The number of pages tall to fit content.</param>
public sealed record PageSetupOptions(
    string? Orientation,
    string? PaperSize,
    double? LeftMargin,
    double? RightMargin,
    double? TopMargin,
    double? BottomMargin,
    string? Header,
    string? Footer,
    bool? FitToPage,
    int? FitToPagesWide,
    int? FitToPagesTall);
