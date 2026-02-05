using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Toc;

/// <summary>
///     Handler for removing table of contents pages from PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemovePdfTocHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes all pages that have TocInfo set from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>Success message indicating how many TOC pages were removed.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;

        var tocPages = new List<int>();
        for (var i = 1; i <= document.Pages.Count; i++)
            if (document.Pages[i].TocInfo != null)
                tocPages.Add(i);

        if (tocPages.Count == 0)
            return new SuccessResult { Message = "No TOC pages found to remove." };

        for (var i = tocPages.Count - 1; i >= 0; i--) document.Pages.Delete(tocPages[i]);

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Removed {tocPages.Count} TOC page(s) from document."
        };
    }
}
