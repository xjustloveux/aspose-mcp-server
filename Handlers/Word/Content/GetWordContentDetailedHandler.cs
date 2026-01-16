using System.Text;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Content;

/// <summary>
///     Handler for getting detailed Word document content including headers and footers.
/// </summary>
public class GetWordContentDetailedHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_content_detailed";

    /// <summary>
    ///     Gets detailed document content including optional headers and footers.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: includeHeaders, includeFooters
    /// </param>
    /// <returns>Detailed document content as plain text.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetContentDetailedParameters(parameters);

        var document = context.Document;
        var sb = new StringBuilder();
        sb.AppendLine("=== Detailed Document Content ===");

        if (p.IncludeHeaders)
        {
            sb.AppendLine("\n--- Headers ---");
            foreach (var section in document.Sections.Cast<Section>())
            foreach (var header in section.HeadersFooters.Cast<Aspose.Words.HeaderFooter>())
                if (header.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                    header.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                    header.HeaderFooterType == HeaderFooterType.HeaderEven)
                {
                    var headerText = WordContentHelper.CleanText(header.GetText());
                    if (!string.IsNullOrWhiteSpace(headerText))
                    {
                        sb.AppendLine($"Section {document.Sections.IndexOf(section)} - {header.HeaderFooterType}:");
                        sb.AppendLine(headerText);
                    }
                }
        }

        sb.AppendLine("\n--- Body Content ---");
        foreach (var section in document.Sections.Cast<Section>())
        {
            var bodyText = WordContentHelper.CleanText(section.Body.GetText());
            if (!string.IsNullOrWhiteSpace(bodyText))
                sb.AppendLine(bodyText);
        }

        if (p.IncludeFooters)
        {
            sb.AppendLine("\n--- Footers ---");
            foreach (var section in document.Sections.Cast<Section>())
            foreach (var footer in section.HeadersFooters.Cast<Aspose.Words.HeaderFooter>())
                if (footer.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                    footer.HeaderFooterType == HeaderFooterType.FooterFirst ||
                    footer.HeaderFooterType == HeaderFooterType.FooterEven)
                {
                    var footerText = WordContentHelper.CleanText(footer.GetText());
                    if (!string.IsNullOrWhiteSpace(footerText))
                    {
                        sb.AppendLine($"Section {document.Sections.IndexOf(section)} - {footer.HeaderFooterType}:");
                        sb.AppendLine(footerText);
                    }
                }
        }

        return sb.ToString();
    }

    /// <summary>
    ///     Extracts parameters for the get content detailed operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetContentDetailedParameters ExtractGetContentDetailedParameters(OperationParameters parameters)
    {
        var includeHeaders = parameters.GetOptional("includeHeaders", false);
        var includeFooters = parameters.GetOptional("includeFooters", false);

        return new GetContentDetailedParameters(includeHeaders, includeFooters);
    }

    /// <summary>
    ///     Parameters for the get content detailed operation.
    /// </summary>
    /// <param name="IncludeHeaders">Whether to include headers in the output.</param>
    /// <param name="IncludeFooters">Whether to include footers in the output.</param>
    private record GetContentDetailedParameters(bool IncludeHeaders, bool IncludeFooters);
}
