namespace AsposeMcpServer.Helpers;

/// <summary>
///     Counts the pixels a render has committed to, page by page, and refuses the next one before
///     its bitmap is allocated.
///     <para>
///         <see cref="RenderBudget.EnsureWithinBudget" /> estimates the whole job up front from an
///         average page size, which is all a caller can do before it knows the geometry. Two things
///         escaped it: a request for a single page was never checked at all — one 200-inch sheet at
///         3,000 DPI is a larger allocation than most whole documents — and a multi-page estimate
///         fell back to A4 whatever the document actually held (R3-R03). This adds up the real
///         dimensions instead, and does it before each page is drawn rather than after.
///     </para>
/// </summary>
public sealed class PixelBudget
{
    private int _pages;
    private double _total;

    /// <summary>Pixels committed so far, saturated at <see cref="long.MaxValue" />.</summary>
    public long Total => (long)Math.Min(_total, long.MaxValue);

    /// <summary>
    ///     Adds one page and refuses the render when the running total passes the limit.
    /// </summary>
    /// <param name="widthInches">The page's real width; A4 is assumed when it is not known.</param>
    /// <param name="heightInches">The page's real height; A4 is assumed when it is not known.</param>
    /// <param name="dpi">Resolution the page will be rendered at.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when this page takes the render past <see cref="RenderBudget.MaxTotalPixels" />.
    /// </exception>
    public void Add(double widthInches, double heightInches, int dpi)
    {
        var width = widthInches > 0 ? widthInches : RenderBudget.DefaultPageWidthInches;
        var height = heightInches > 0 ? heightInches : RenderBudget.DefaultPageHeightInches;
        var resolution = dpi > 0 ? dpi : 96;

        // Kept in double on purpose: the product of a large sheet and a high resolution overflows
        // a long, and an overflowed total compares as smaller than the limit rather than larger.
        _total += width * resolution * (height * resolution);
        _pages++;

        if (_total <= RenderBudget.MaxTotalPixels) return;

        throw new ArgumentException(
            $"Rendering {_pages:N0} page(s) at {resolution} DPI comes to roughly {Total:N0} pixels, "
            + $"above the limit of {RenderBudget.MaxTotalPixels:N0}. Lower the DPI, or render fewer "
            + "pages at a time.");
    }
}
