using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Text;

/// <summary>
///     Handler for extracting text from PDF documents.
/// </summary>
public class ExtractPdfTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>
    ///     Extracts text from the specified page of the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex, includeFontInfo, extractionMode
    /// </param>
    /// <returns>JSON string containing the extracted text.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractExtractParameters(parameters);

        var document = context.Document;
        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];

        var textAbsorber = new TextAbsorber();
        if (string.Equals(p.ExtractionMode, "raw", StringComparison.OrdinalIgnoreCase))
            textAbsorber.ExtractionOptions = new TextExtractionOptions(TextExtractionOptions.TextFormattingMode.Raw);

        page.Accept(textAbsorber);

        if (p.IncludeFontInfo)
        {
            var textFragmentAbsorber = new TextFragmentAbsorber();
            page.Accept(textFragmentAbsorber);
            List<object> fragments = [];

            foreach (var fragment in textFragmentAbsorber.TextFragments)
                fragments.Add(new
                {
                    text = fragment.Text,
                    fontName = fragment.TextState.Font.FontName,
                    fontSize = fragment.TextState.FontSize
                });

            return JsonResult(new
            {
                pageIndex = p.PageIndex,
                totalPages = document.Pages.Count,
                fragmentCount = fragments.Count,
                fragments
            });
        }

        return JsonResult(new
        {
            pageIndex = p.PageIndex,
            totalPages = document.Pages.Count,
            text = textAbsorber.Text
        });
    }

    /// <summary>
    ///     Extracts parameters for extract operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static ExtractParameters ExtractExtractParameters(OperationParameters parameters)
    {
        return new ExtractParameters(
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional("includeFontInfo", false),
            parameters.GetOptional("extractionMode", "pure")
        );
    }

    /// <summary>
    ///     Parameters for extract operation.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="IncludeFontInfo">Whether to include font information.</param>
    /// <param name="ExtractionMode">The text extraction mode.</param>
    private record ExtractParameters(int PageIndex, bool IncludeFontInfo, string ExtractionMode);
}
