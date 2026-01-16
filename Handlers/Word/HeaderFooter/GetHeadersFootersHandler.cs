using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for getting headers and footers from Word documents.
/// </summary>
public class GetHeadersFootersHandler : OperationHandlerBase<Document>
{
    private static readonly (HeaderFooterType type, string name)[] HeaderTypes =
    [
        (HeaderFooterType.HeaderPrimary, "primary"),
        (HeaderFooterType.HeaderFirst, "firstPage"),
        (HeaderFooterType.HeaderEven, "evenPage")
    ];

    private static readonly (HeaderFooterType type, string name)[] FooterTypes =
    [
        (HeaderFooterType.FooterPrimary, "primary"),
        (HeaderFooterType.FooterFirst, "firstPage"),
        (HeaderFooterType.FooterEven, "evenPage")
    ];

    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all headers and footers from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sectionIndex (specific section or -1 for all)
    /// </param>
    /// <returns>A JSON string containing headers and footers information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetHeadersFootersParameters(parameters);
        var doc = context.Document;
        doc.UpdateFields();

        ValidateSectionIndex(doc, p.SectionIndex);
        var sections = GetTargetSections(doc, p.SectionIndex);

        var sectionsList = sections.Select((section, i) =>
            BuildSectionInfo(section, GetActualIndex(p.SectionIndex, i))).ToList();

        return JsonResult(new
        {
            totalSections = doc.Sections.Count,
            queriedSectionIndex = p.SectionIndex,
            sections = sectionsList
        });
    }

    /// <summary>
    ///     Validates that the section index is within range.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="sectionIndex">The section index to validate.</param>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    private static void ValidateSectionIndex(Document doc, int? sectionIndex)
    {
        if (sectionIndex.HasValue && sectionIndex.Value != -1 &&
            (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count))
            throw new ArgumentException(
                $"Section index {sectionIndex.Value} is out of range (document has {doc.Sections.Count} sections)");
    }

    /// <summary>
    ///     Gets the target sections based on section index.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="sectionIndex">The section index, or null/-1 for all sections.</param>
    /// <returns>An array of target sections.</returns>
    private static Section[] GetTargetSections(Document doc, int? sectionIndex)
    {
        return sectionIndex.HasValue && sectionIndex.Value != -1
            ? [doc.Sections[sectionIndex.Value]]
            : doc.Sections.Cast<Section>().ToArray();
    }

    /// <summary>
    ///     Gets the actual section index for the loop iteration.
    /// </summary>
    /// <param name="sectionIndex">The requested section index.</param>
    /// <param name="loopIndex">The current loop index.</param>
    /// <returns>The actual section index to use.</returns>
    private static int GetActualIndex(int? sectionIndex, int loopIndex)
    {
        return sectionIndex.HasValue && sectionIndex.Value != -1 ? sectionIndex.Value : loopIndex;
    }

    /// <summary>
    ///     Builds the section information object including headers and footers.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <param name="actualIndex">The actual section index.</param>
    /// <returns>An object containing the section information.</returns>
    private static object BuildSectionInfo(Section section, int actualIndex)
    {
        var headers = ExtractHeaderFooterText(section, HeaderTypes);
        var footers = ExtractHeaderFooterText(section, FooterTypes);

        return new
        {
            sectionIndex = actualIndex,
            headers = headers.Count > 0 ? headers : null,
            footers = footers.Count > 0 ? footers : null
        };
    }

    /// <summary>
    ///     Extracts text content from headers or footers.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <param name="types">The header/footer types to extract.</param>
    /// <returns>A dictionary of header/footer names to their text content.</returns>
    private static Dictionary<string, string?> ExtractHeaderFooterText(Section section,
        (HeaderFooterType type, string name)[] types)
    {
        var result = new Dictionary<string, string?>();
        foreach (var (type, name) in types)
        {
            var hf = section.HeadersFooters[type];
            if (hf == null) continue;

            var text = hf.ToString(SaveFormat.Text).Trim();
            if (!string.IsNullOrEmpty(text))
                result[name] = text;
        }

        return result;
    }

    /// <summary>
    ///     Extracts parameters for the get headers footers operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetHeadersFootersParameters ExtractGetHeadersFootersParameters(OperationParameters parameters)
    {
        return new GetHeadersFootersParameters(
            parameters.GetOptional<int?>("sectionIndex")
        );
    }

    /// <summary>
    ///     Parameters for the get headers footers operation.
    /// </summary>
    /// <param name="SectionIndex">The section index, or null/-1 for all sections.</param>
    private sealed record GetHeadersFootersParameters(int? SectionIndex);
}
