using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for adding pages to PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddPdfPageHandler : OperationHandlerBase<Document>
{
    /// <summary>
    ///     Largest number of pages one call may add. Ten thousand empty pages is not a document
    ///     anyone is authoring; it is a memory and file-size request, and the previous cap was
    ///     high enough to be no cap at all in practice (R2-R06). A caller that genuinely needs
    ///     more can issue more calls, which keeps each one bounded.
    /// </summary>
    private const int MaxPagesPerCall = 1_000;

    /// <summary>
    ///     Largest page dimension in points. The PDF format itself stops at 14,400 points
    ///     (200 inches); anything beyond that cannot be represented, and a page approaching it is
    ///     already a rendering cost rather than a document (R2-R06).
    /// </summary>
    private const double MaxPageDimensionPoints = 14_400;

    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds one or more pages to the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: count (number of pages to add, default: 1),
    ///     insertAt (1-based position to insert pages at),
    ///     width (page width in points),
    ///     height (page height in points)
    /// </param>
    /// <returns>Success message with new page information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);
        // Both bounds. Leaving the lower one at int.MinValue let count=0 or -1 through: the
        // loop added nothing, the document was still marked modified, and the response claimed
        // "Added -1 page(s)" (R3-C04).
        SecurityHelper.ValidateNumericRange(p.Count, "count", 1, MaxPagesPerCall);

        // Validated before the loop: SetPageSize ran after the page had been added, so a refused
        // size left a blank page behind and the document was reported unmodified while carrying it
        // (R4-R05).
        EnsurePageSizeUsable(p.Width, p.Height);

        var doc = context.Document;
        var shouldInsert = p.InsertAt is >= 1 && p.InsertAt.Value <= doc.Pages.Count;

        for (var i = 0; i < p.Count; i++)
        {
            var page = shouldInsert ? doc.Pages.Insert(p.InsertAt!.Value + i) : doc.Pages.Add();
            SetPageSize(page, p.Width, p.Height);
        }

        MarkModified(context);

        return new SuccessResult { Message = $"Added {p.Count} page(s). Total pages: {doc.Pages.Count}" };
    }

    /// <summary>
    ///     Extracts add parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetOptional("count", 1),
            parameters.GetOptional<int?>("insertAt"),
            parameters.GetOptional<double?>("width"),
            parameters.GetOptional<double?>("height"));
    }

    /// <summary>
    ///     Refuses a page size before any page exists to apply it to.
    ///     <para>
    ///         Each dimension is checked on its own. Skipping both whenever either was absent meant
    ///         a caller who gave only a width had it neither validated nor used (R5-C01).
    ///     </para>
    /// </summary>
    /// <param name="width">Requested width in points, or <c>null</c> to keep A4's.</param>
    /// <param name="height">Requested height in points, or <c>null</c> to keep A4's.</param>
    /// <exception cref="ArgumentException">Thrown when either dimension is unusable.</exception>
    private static void EnsurePageSizeUsable(double? width, double? height)
    {
        if (width.HasValue) EnsureUsableDimension(width.Value, "width");
        if (height.HasValue) EnsureUsableDimension(height.Value, "height");
    }

    /// <summary>
    ///     Sets the page size, falling back to A4 for whichever dimension the caller left out.
    ///     <para>
    ///         A missing dimension used to discard the one that was given: a width-only request
    ///         produced a plain A4 page and still reported success, so the caller had no way to
    ///         tell their width had been ignored (R5-C01).
    ///     </para>
    /// </summary>
    /// <param name="page">The page to set the size for.</param>
    /// <param name="width">The optional width in points.</param>
    /// <param name="height">The optional height in points.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when either dimension is not a usable measurement.
    /// </exception>
    private static void SetPageSize(Aspose.Pdf.Page page, double? width, double? height)
    {
        EnsurePageSizeUsable(width, height);
        page.SetPageSize(width ?? PageSize.A4.Width, height ?? PageSize.A4.Height);
    }

    /// <summary>
    ///     Refuses a page dimension that is not a finite, positive, representable measurement.
    /// </summary>
    /// <param name="value">The requested dimension in points.</param>
    /// <param name="name">Parameter name for the error message.</param>
    /// <exception cref="ArgumentException">Thrown when the value cannot describe a page.</exception>
    private static void EnsureUsableDimension(double value, string name)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
            throw new ArgumentException($"{name} must be a finite number of points.", name);
        if (value <= 0)
            throw new ArgumentException($"{name} must be greater than zero.", name);
        if (value > MaxPageDimensionPoints)
            throw new ArgumentException(
                $"{name} of {value:N0} points exceeds the PDF maximum of "
                + $"{MaxPageDimensionPoints:N0} points (200 inches).", name);
    }

    /// <summary>
    ///     Parameters for adding pages.
    /// </summary>
    /// <param name="Count">The number of pages to add.</param>
    /// <param name="InsertAt">The optional 1-based position to insert pages at.</param>
    /// <param name="Width">The optional page width in points.</param>
    /// <param name="Height">The optional page height in points.</param>
    private sealed record AddParameters(int Count, int? InsertAt, double? Width, double? Height);
}
