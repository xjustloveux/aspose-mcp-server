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
        var p = ExtractGetContentParameters(parameters);

        var document = context.Document;

        if (p.PageIndex.HasValue)
        {
            if (p.PageIndex.Value < 1 || p.PageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var textAbsorber = new TextAbsorber();
            document.Pages[p.PageIndex.Value].Accept(textAbsorber);

            return JsonResult(new
            {
                pageIndex = p.PageIndex.Value,
                totalPages = document.Pages.Count,
                content = textAbsorber.Text
            });
        }

        var pagesToExtract = Math.Min(p.MaxPages, document.Pages.Count);
        var truncated = document.Pages.Count > p.MaxPages;
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

    /// <summary>
    ///     Extracts get content parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetContentParameters ExtractGetContentParameters(OperationParameters parameters)
    {
        return new GetContentParameters(
            parameters.GetOptional<int?>("pageIndex"),
            parameters.GetOptional("maxPages", 100));
    }

    /// <summary>
    ///     Parameters for getting content.
    /// </summary>
    /// <param name="PageIndex">The optional 1-based page index.</param>
    /// <param name="MaxPages">The maximum number of pages to extract.</param>
    private record GetContentParameters(int? PageIndex, int MaxPages);
}
