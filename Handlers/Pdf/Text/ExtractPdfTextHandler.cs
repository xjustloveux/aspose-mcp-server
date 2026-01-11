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
        var pageIndex = parameters.GetOptional("pageIndex", 1);
        var includeFontInfo = parameters.GetOptional("includeFontInfo", false);
        var extractionMode = parameters.GetOptional("extractionMode", "pure");

        var document = context.Document;
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        var textAbsorber = new TextAbsorber();
        if (extractionMode.ToLower() == "raw")
            textAbsorber.ExtractionOptions = new TextExtractionOptions(TextExtractionOptions.TextFormattingMode.Raw);

        page.Accept(textAbsorber);

        if (includeFontInfo)
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
                pageIndex,
                totalPages = document.Pages.Count,
                fragmentCount = fragments.Count,
                fragments
            });
        }

        return JsonResult(new
        {
            pageIndex,
            totalPages = document.Pages.Count,
            text = textAbsorber.Text
        });
    }
}
