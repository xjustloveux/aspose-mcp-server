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
            AppendHeadersFooters(sb, document, "Headers", IsHeaderType);

        AppendBodyContent(sb, document);

        if (p.IncludeFooters)
            AppendHeadersFooters(sb, document, "Footers", IsFooterType);

        return sb.ToString();
    }

    /// <summary>
    ///     Appends headers or footers content to the StringBuilder.
    /// </summary>
    private static void AppendHeadersFooters(StringBuilder sb, Document document, string sectionName,
        Func<HeaderFooterType, bool> typeFilter)
    {
        sb.AppendLine($"\n--- {sectionName} ---");
        foreach (var section in document.Sections.Cast<Section>())
        foreach (var hf in section.HeadersFooters.Cast<Aspose.Words.HeaderFooter>())
        {
            if (!typeFilter(hf.HeaderFooterType)) continue;

            var text = WordContentHelper.CleanText(hf.GetText());
            if (string.IsNullOrWhiteSpace(text)) continue;

            sb.AppendLine($"Section {document.Sections.IndexOf(section)} - {hf.HeaderFooterType}:");
            sb.AppendLine(text);
        }
    }

    /// <summary>
    ///     Appends body content to the StringBuilder.
    /// </summary>
    private static void AppendBodyContent(StringBuilder sb, Document document)
    {
        sb.AppendLine("\n--- Body Content ---");
        foreach (var section in document.Sections.Cast<Section>())
        {
            var bodyText = WordContentHelper.CleanText(section.Body.GetText());
            if (!string.IsNullOrWhiteSpace(bodyText))
                sb.AppendLine(bodyText);
        }
    }

    /// <summary>
    ///     Determines if the header/footer type is a header type.
    /// </summary>
    private static bool IsHeaderType(HeaderFooterType type)
    {
        return type == HeaderFooterType.HeaderPrimary ||
               type == HeaderFooterType.HeaderFirst ||
               type == HeaderFooterType.HeaderEven;
    }

    /// <summary>
    ///     Determines if the header/footer type is a footer type.
    /// </summary>
    private static bool IsFooterType(HeaderFooterType type)
    {
        return type == HeaderFooterType.FooterPrimary ||
               type == HeaderFooterType.FooterFirst ||
               type == HeaderFooterType.FooterEven;
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
    private sealed record GetContentDetailedParameters(bool IncludeHeaders, bool IncludeFooters);
}
