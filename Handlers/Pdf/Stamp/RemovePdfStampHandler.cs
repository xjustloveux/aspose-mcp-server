using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Stamp;

/// <summary>
///     Handler for removing stamp annotations from PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemovePdfStampHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes one or all stamp annotations from the specified page.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex (1-based).
    ///     Optional: stampIndex (1-based, if omitted removes all stamp annotations on the page).
    /// </param>
    /// <returns>Success message with removal details.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the page index is out of range or the stamp index is invalid.
    /// </exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        var document = context.Document;

        AddTextPdfStampHandler.ValidatePageIndex(p.PageIndex, document);
        var page = document.Pages[p.PageIndex];

        if (p.StampIndex.HasValue)
        {
            if (p.StampIndex.Value < 1 || p.StampIndex.Value > page.Annotations.Count)
                throw new ArgumentException($"stampIndex must be between 1 and {page.Annotations.Count}");

            page.Annotations.Delete(p.StampIndex.Value);

            MarkModified(context);

            return new SuccessResult
            {
                Message = $"Removed stamp annotation {p.StampIndex.Value} from page {p.PageIndex}."
            };
        }

        var stampsToRemove = new List<Aspose.Pdf.Annotations.Annotation>();
        foreach (var annotation in page.Annotations)
            if (annotation is StampAnnotation)
                stampsToRemove.Add(annotation);

        if (stampsToRemove.Count == 0)
            return new SuccessResult { Message = $"No stamp annotations found on page {p.PageIndex}." };

        foreach (var s in stampsToRemove)
            page.Annotations.Delete(s);

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Removed all {stampsToRemove.Count} stamp annotation(s) from page {p.PageIndex}."
        };
    }

    /// <summary>
    ///     Extracts parameters for the remove stamp operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static RemoveStampParameters ExtractParameters(OperationParameters parameters)
    {
        return new RemoveStampParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetOptional<int?>("stampIndex")
        );
    }

    /// <summary>
    ///     Parameters for the remove stamp operation.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="StampIndex">The 1-based stamp annotation index (null = remove all stamps on the page).</param>
    private sealed record RemoveStampParameters(int PageIndex, int? StampIndex);
}
