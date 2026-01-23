using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Annotation;

/// <summary>
///     Handler for deleting annotations from PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteParameters(parameters);

        var document = context.Document;

        if (deleteParams.PageIndex < 1 || deleteParams.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[deleteParams.PageIndex];

        if (deleteParams.AnnotationIndex.HasValue)
        {
            if (deleteParams.AnnotationIndex.Value < 1 || deleteParams.AnnotationIndex.Value > page.Annotations.Count)
                throw new ArgumentException($"annotationIndex must be between 1 and {page.Annotations.Count}");

            page.Annotations.Delete(deleteParams.AnnotationIndex.Value);

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Deleted annotation {deleteParams.AnnotationIndex.Value} from page {deleteParams.PageIndex}."
            };
        }

        var count = page.Annotations.Count;
        if (count == 0)
            throw new ArgumentException($"No annotations found on page {deleteParams.PageIndex}");

        page.Annotations.Delete();

        MarkModified(context);

        return new SuccessResult { Message = $"Deleted all {count} annotation(s) from page {deleteParams.PageIndex}." };
    }

    /// <summary>
    ///     Extracts delete parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetOptional<int?>("annotationIndex")
        );
    }

    /// <summary>
    ///     Record to hold delete annotation parameters.
    /// </summary>
    private sealed record DeleteParameters(int PageIndex, int? AnnotationIndex);
}
