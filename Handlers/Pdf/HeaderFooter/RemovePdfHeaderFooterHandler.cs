using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Pdf;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.HeaderFooter;

/// <summary>
///     Handler for removing stamp annotations from PDF document pages.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemovePdfHeaderFooterHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes stamp annotations from specified pages of the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageRange (if null, removes from all pages).
    /// </param>
    /// <returns>Success message with the number of stamps removed.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageRange = parameters.GetOptional<string?>("pageRange");

        var document = context.Document;

        var pages = PdfWatermarkHelper.ParsePageRange(pageRange, document.Pages.Count);
        var removedCount = 0;

        foreach (var pageIdx in pages)
        {
            var page = document.Pages[pageIdx];

            var stampsToRemove = new List<Aspose.Pdf.Annotations.Annotation>();
            foreach (var annotation in page.Annotations)
                if (annotation is StampAnnotation)
                    stampsToRemove.Add(annotation);

            foreach (var stamp in stampsToRemove)
            {
                page.Annotations.Delete(stamp);
                removedCount++;
            }
        }

        if (removedCount > 0)
            MarkModified(context);

        return new SuccessResult
        {
            Message = $"Removed {removedCount} stamp(s) from {pages.Count} page(s)."
        };
    }
}
