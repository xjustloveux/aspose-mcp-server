using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Link;

/// <summary>
///     Handler for adding links to PDF documents.
/// </summary>
public class AddPdfLinkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a link to a specific page in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex, url or targetPage.
    ///     Optional: x, y, width, height.
    /// </param>
    /// <returns>Success message with link details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);

        if (string.IsNullOrEmpty(p.Url) && !p.TargetPage.HasValue)
            throw new ArgumentException("Either 'url' or 'targetPage' is required");

        var document = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];
        var rect = new Rectangle(p.X, p.Y, p.X + p.Width, p.Y + p.Height);

        var linkAnnotation = new LinkAnnotation(page, rect);

        if (!string.IsNullOrEmpty(p.Url))
        {
            linkAnnotation.Action = new GoToURIAction(p.Url);
        }
        else if (p.TargetPage.HasValue)
        {
            if (p.TargetPage.Value < 1 || p.TargetPage.Value > document.Pages.Count)
                throw new ArgumentException($"targetPage must be between 1 and {document.Pages.Count}");

            linkAnnotation.Action = new GoToAction(document.Pages[p.TargetPage.Value]);
        }

        page.Annotations.Add(linkAnnotation);

        MarkModified(context);

        var linkType = !string.IsNullOrEmpty(p.Url) ? $"URL: {p.Url}" : $"Page: {p.TargetPage}";
        return Success($"Link added to page {p.PageIndex} ({linkType}).");
    }

    /// <summary>
    ///     Extracts add parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetOptional<string?>("url"),
            parameters.GetOptional<int?>("targetPage"),
            parameters.GetOptional("x", 100.0),
            parameters.GetOptional("y", 700.0),
            parameters.GetOptional("width", 100.0),
            parameters.GetOptional("height", 20.0));
    }

    /// <summary>
    ///     Parameters for adding a link.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="Url">The optional URL for external links.</param>
    /// <param name="TargetPage">The optional target page for internal links.</param>
    /// <param name="X">The X coordinate.</param>
    /// <param name="Y">The Y coordinate.</param>
    /// <param name="Width">The width of the link area.</param>
    /// <param name="Height">The height of the link area.</param>
    private sealed record AddParameters(
        int PageIndex,
        string? Url,
        int? TargetPage,
        double X,
        double Y,
        double Width,
        double Height);
}
