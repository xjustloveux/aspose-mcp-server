using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Annotation;

/// <summary>
///     Handler for deleting annotations from PDF documents.
/// </summary>
public class DeletePdfAnnotationHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes one or all annotations from the specified page.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex
    ///     Optional: annotationIndex (if omitted, deletes all annotations on the page)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetRequired<int>("pageIndex");
        var annotationIndex = parameters.GetOptional<int?>("annotationIndex");

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        if (annotationIndex.HasValue)
        {
            if (annotationIndex.Value < 1 || annotationIndex.Value > page.Annotations.Count)
                throw new ArgumentException($"annotationIndex must be between 1 and {page.Annotations.Count}");

            page.Annotations.Delete(annotationIndex.Value);

            MarkModified(context);

            return Success($"Deleted annotation {annotationIndex.Value} from page {pageIndex}.");
        }

        var count = page.Annotations.Count;
        if (count == 0)
            throw new ArgumentException($"No annotations found on page {pageIndex}");

        page.Annotations.Delete();

        MarkModified(context);

        return Success($"Deleted all {count} annotation(s) from page {pageIndex}.");
    }
}
