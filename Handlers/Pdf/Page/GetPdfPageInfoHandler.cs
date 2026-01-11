using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for getting information about all pages in PDF documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        List<object> pageList = [];

        for (var i = 1; i <= doc.Pages.Count; i++)
        {
            var page = doc.Pages[i];
            pageList.Add(new
            {
                pageIndex = i,
                width = page.Rect.Width,
                height = page.Rect.Height,
                rotation = page.Rotate.ToString()
            });
        }

        var result = new
        {
            count = doc.Pages.Count,
            items = pageList
        };

        return JsonResult(result);
    }
}
