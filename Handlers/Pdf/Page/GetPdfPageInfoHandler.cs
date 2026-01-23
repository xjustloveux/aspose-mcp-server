using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Page;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for getting information about all pages in PDF documents.
/// </summary>
[ResultType(typeof(GetPdfPageInfoResult))]
public class GetPdfPageInfoHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_info";

    /// <summary>
    ///     Gets basic information about all pages in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing page information for all pages.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        List<PdfPageInfo> pageList = [];

        for (var i = 1; i <= doc.Pages.Count; i++)
        {
            var page = doc.Pages[i];
            pageList.Add(new PdfPageInfo
            {
                PageIndex = i,
                Width = page.Rect.Width,
                Height = page.Rect.Height,
                Rotation = page.Rotate.ToString()
            });
        }

        return new GetPdfPageInfoResult
        {
            Count = doc.Pages.Count,
            Items = pageList
        };
    }
}
