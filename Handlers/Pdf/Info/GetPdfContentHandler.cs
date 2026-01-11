using System.Text;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Info;

/// <summary>
///     Handler for extracting text content from PDF documents.
/// </summary>
public class GetPdfContentHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_content";

    /// <summary>
    ///     Extracts text content from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex (1-based), maxPages (default: 100).
    /// </param>
    /// <returns>JSON string containing the extracted text content.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetOptional<int?>("pageIndex");
        var maxPages = parameters.GetOptional("maxPages", 100);

        var document = context.Document;

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var textAbsorber = new TextAbsorber();
            document.Pages[pageIndex.Value].Accept(textAbsorber);

            return JsonResult(new
            {
                pageIndex = pageIndex.Value,
                totalPages = document.Pages.Count,
                content = textAbsorber.Text
            });
        }

        var pagesToExtract = Math.Min(maxPages, document.Pages.Count);
        var truncated = document.Pages.Count > maxPages;
        var contentBuilder = new StringBuilder();

        for (var i = 1; i <= pagesToExtract; i++)
        {
            var textAbsorber = new TextAbsorber();
            document.Pages[i].Accept(textAbsorber);
            contentBuilder.AppendLine(textAbsorber.Text);
        }

        return JsonResult(new
        {
            totalPages = document.Pages.Count,
            extractedPages = pagesToExtract,
            truncated,
            content = contentBuilder.ToString()
        });
    }
}
