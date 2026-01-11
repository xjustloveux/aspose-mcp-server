using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Annotation;

/// <summary>
///     Handler for adding text annotations to PDF documents.
/// </summary>
public class AddPdfAnnotationHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new text annotation to the specified page.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex, text
    ///     Optional: x (default: 100), y (default: 700)
    /// </param>
    /// <returns>Success message with annotation details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetRequired<int>("pageIndex");
        var text = parameters.GetRequired<string>("text");
        var x = parameters.GetOptional("x", 100.0);
        var y = parameters.GetOptional("y", 700.0);

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var annotation = new TextAnnotation(page, new Rectangle(x, y, x + 200, y + 50))
        {
            Title = "Comment",
            Subject = "Annotation",
            Contents = text,
            Open = false,
            Icon = TextIcon.Note
        };

        page.Annotations.Add(annotation);

        MarkModified(context);

        return Success($"Added annotation to page {pageIndex}.");
    }
}
