using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting header tab stops in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetHeaderTabsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_header_tabs";

    /// <summary>
    ///     Sets tab stops in the document header.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tabStops (array), sectionIndex, headerFooterType
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetHeaderTabsParameters(parameters);

        var doc = context.Document;
        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType, true);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        // Read once, before any section is touched: resolving per section would let a malformed
        // entry throw after earlier sections had already had their stops replaced.
        var resolved = WordTabStopHelper.Resolve(p.TabStops);

        foreach (var section in sections)
        {
            var header = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);
            if (resolved.Count > 0) ApplyTabStops(doc, header, resolved);
        }

        MarkModified(context);
        return new SuccessResult { Message = "Header tab stops set" };
    }

    /// <summary>
    ///     Applies already-resolved tab stops to the header paragraph.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="header">The header.</param>
    /// <param name="resolved">
    ///     The stops to apply, read by <see cref="WordTabStopHelper" /> before this was called.
    ///     Reading each entry as it was applied meant a malformed one threw after Clear(), leaving
    ///     the paragraph with no stops at all — the defect fixed for `edit` in §23.4 and left
    ///     standing here (R13-W01).
    /// </param>
    private static void ApplyTabStops(
        Document doc, Aspose.Words.HeaderFooter header, List<TabStop> resolved)
    {
        var para = header.FirstParagraph ?? new WordParagraph(doc);
        para.ParagraphFormat.TabStops.Clear();
        foreach (var stop in resolved) para.ParagraphFormat.TabStops.Add(stop);

        if (header.FirstParagraph == null) header.AppendChild(para);
    }

    /// <summary>
    ///     Extracts parameters for the set header tabs operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetHeaderTabsParameters ExtractSetHeaderTabsParameters(OperationParameters parameters)
    {
        return new SetHeaderTabsParameters(
            parameters.GetOptional<JsonArray?>("tabStops"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional("headerFooterType", "primary")
        );
    }

    /// <summary>
    ///     Parameters for the set header tabs operation.
    /// </summary>
    /// <param name="TabStops">The tab stops array.</param>
    /// <param name="SectionIndex">The section index.</param>
    /// <param name="HeaderFooterType">The header/footer type.</param>
    private sealed record SetHeaderTabsParameters(
        JsonArray? TabStops,
        int SectionIndex,
        string HeaderFooterType);
}
