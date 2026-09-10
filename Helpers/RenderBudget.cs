namespace AsposeMcpServer.Helpers;

/// <summary>
///     The aggregate cost of a rendering request, checked before the work starts (A-06).
///     <para>
///         Individual parameters were capped one at a time — DPI to 1200, page counts, array
///         sizes — but a request's real cost is their product. A 500-page document at 1200 DPI is
///         within every single limit and still asks for roughly a hundred billion pixels, so the
///         only meaningful bound is on the total.
///     </para>
///     <para>
///         The estimate uses page area at the requested resolution. It does not have to be exact:
///         it exists to separate "a document someone wants to look at" from "a request that will
///         occupy this process for minutes and exhaust its memory", and those differ by orders of
///         magnitude.
///     </para>
/// </summary>
public static class RenderBudget
{
    /// <summary>
    ///     Largest total pixel count one request may render.
    ///     <para>
    ///         Two billion is about 250 A4 pages at 300 DPI, or 40 at 600 — comfortably more than
    ///         any real conversion, and far below the point where the allocations stop being
    ///         serviceable.
    ///     </para>
    /// </summary>
    public const long MaxTotalPixels = 2_000_000_000;

    /// <summary>Largest number of image files one request may produce.</summary>
    public const int MaxOutputFiles = 5_000;

    /// <summary>
    ///     Largest total size one request may write out.
    ///     <para>
    ///         A page count says nothing about what lands on disk: one worksheet can raster to
    ///         hundreds of megabytes, and an archive can hold thousands of small parts whose sum is
    ///         larger than either limit alone would catch (R2-R02).
    ///     </para>
    /// </summary>
    public const long MaxOutputBytes = 2L * 1024 * 1024 * 1024;

    /// <summary>
    ///     Most bytes a conversion may produce when the result is held in memory rather than
    ///     written to a file.
    ///     <para>
    ///         <see cref="MaxOutputBytes" /> is the right bound for a file: it is written once and
    ///         streamed. A response that is built in memory and then handed back as a byte array is
    ///         held at least twice over at the moment of the copy, and several sessions converting
    ///         at once multiply that — so a request approaching the on-disk limit could hold several
    ///         gigabytes of managed memory (§17.4.2). A caller that needs more than this converts to
    ///         a file instead.
    ///     </para>
    /// </summary>
    public const long MaxInMemoryOutputBytes = 64L * 1024 * 1024;

    /// <summary>
    ///     A4 portrait, used when the real page size is not available.
    /// </summary>
    public const double DefaultPageWidthInches = 8.27;

    /// <inheritdoc cref="DefaultPageWidthInches" />
    public const double DefaultPageHeightInches = 11.69;

    /// <summary>
    ///     Refuses a request that would write out more files than the budget allows.
    /// </summary>
    /// <param name="fileCount">Files the request would produce.</param>
    /// <param name="what">What is being written, for the message.</param>
    /// <exception cref="ArgumentException">Thrown when the count exceeds the budget.</exception>
    public static void EnsureOutputCount(int fileCount, string what = "files")
    {
        if (fileCount > MaxOutputFiles)
            throw new ArgumentException(
                $"The request would produce {fileCount:N0} {what}, above the limit of "
                + $"{MaxOutputFiles:N0}. Extract a subset instead of the whole document.");
    }

    /// <summary>
    ///     Refuses a request once what it has written exceeds the budget.
    ///     <para>
    ///         Called as the work proceeds rather than only up front, because the size of each
    ///         part is only known once it has been produced.
    ///     </para>
    /// </summary>
    /// <param name="bytesSoFar">Bytes written so far.</param>
    /// <param name="what">What is being written, for the message.</param>
    /// <exception cref="ArgumentException">Thrown when the running total exceeds the budget.</exception>
    public static void EnsureOutputBytes(long bytesSoFar, string what = "output")
    {
        if (bytesSoFar > MaxOutputBytes)
            throw new ArgumentException(
                $"The request has produced {bytesSoFar:N0} bytes of {what}, above the limit of "
                + $"{MaxOutputBytes:N0}. Extract a subset instead of the whole document.");
    }

    /// <summary>
    ///     Refuses a render whose estimated total size exceeds the budget.
    /// </summary>
    /// <param name="pageCount">Pages, sheets or slides that would be rendered.</param>
    /// <param name="dpi">Resolution the caller asked for.</param>
    /// <param name="averagePageWidthInches">Average page width; A4 portrait when unknown.</param>
    /// <param name="averagePageHeightInches">Average page height; A4 portrait when unknown.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the request would produce too many files or too many pixels.
    /// </exception>
    public static void EnsureWithinBudget(int pageCount, int dpi,
        double averagePageWidthInches = DefaultPageWidthInches,
        double averagePageHeightInches = DefaultPageHeightInches)
    {
        if (pageCount <= 0) return;

        if (pageCount > MaxOutputFiles)
            throw new ArgumentException(
                $"The request would produce {pageCount:N0} image files, above the limit of "
                + $"{MaxOutputFiles:N0}. Convert a page range instead of the whole document.");

        // A page whose real size is unknown falls back to A4; a caller that knows the actual
        // dimensions passes them, because a 200-inch sheet and an A4 page cost very different
        // amounts at the same DPI (R2-R01).
        var width = averagePageWidthInches > 0 ? averagePageWidthInches : DefaultPageWidthInches;
        var height = averagePageHeightInches > 0 ? averagePageHeightInches : DefaultPageHeightInches;
        var pixelsPerPage = width * dpi * (height * dpi);
        var total = (long)Math.Min(pixelsPerPage * pageCount, long.MaxValue);

        if (total > MaxTotalPixels)
            throw new ArgumentException(
                $"Rendering {pageCount:N0} page(s) at {dpi} DPI comes to roughly {total:N0} pixels, "
                + $"above the limit of {MaxTotalPixels:N0}. Lower the DPI, or convert fewer pages "
                + "at a time.");
    }
}
