namespace AsposeMcpServer.Helpers.Excel;

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
