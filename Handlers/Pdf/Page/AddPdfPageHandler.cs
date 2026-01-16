using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for adding pages to PDF documents.
/// </summary>
public class AddPdfPageHandler : OperationHandlerBase<Document>
{
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);

        var doc = context.Document;
        var shouldInsert = p.InsertAt is >= 1 && p.InsertAt.Value <= doc.Pages.Count;

        for (var i = 0; i < p.Count; i++)
        {
            var page = shouldInsert ? doc.Pages.Insert(p.InsertAt!.Value + i) : doc.Pages.Add();
            SetPageSize(page, p.Width, p.Height);
        }

        MarkModified(context);

        return Success($"Added {p.Count} page(s). Total pages: {doc.Pages.Count}");
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
    ///     Sets the page size with the specified dimensions or defaults to A4.
    /// </summary>
    /// <param name="page">The page to set the size for.</param>
    /// <param name="width">The optional width in points.</param>
    /// <param name="height">The optional height in points.</param>
    private static void SetPageSize(Aspose.Pdf.Page page, double? width, double? height)
    {
        if (width.HasValue && height.HasValue)
            page.SetPageSize(width.Value, height.Value);
        else
            page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
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
