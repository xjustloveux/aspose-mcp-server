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
        var count = parameters.GetOptional("count", 1);
        var insertAt = parameters.GetOptional<int?>("insertAt");
        var width = parameters.GetOptional<double?>("width");
        var height = parameters.GetOptional<double?>("height");

        var doc = context.Document;

        if (insertAt is >= 1 && insertAt.Value <= doc.Pages.Count)
            for (var i = 0; i < count; i++)
            {
                var page = doc.Pages.Insert(insertAt.Value + i);
                if (width.HasValue && height.HasValue)
                    page.SetPageSize(width.Value, height.Value);
                else
                    page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
            }
        else
            for (var i = 0; i < count; i++)
            {
                var page = doc.Pages.Add();
                if (width.HasValue && height.HasValue)
                    page.SetPageSize(width.Value, height.Value);
                else
                    page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
            }

        MarkModified(context);

        return Success($"Added {count} page(s). Total pages: {doc.Pages.Count}");
    }
}
