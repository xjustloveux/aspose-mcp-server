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
///     Handler for setting footer tab stops in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetFooterTabsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_footer_tabs";

    /// <summary>
    ///     Sets tab stops in the document footer.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tabStops (array), sectionIndex, headerFooterType
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetFooterTabsParameters(parameters);

        var doc = context.Document;
        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType, false);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        // Read once, before any section is touched: resolving per section would let a malformed
        // entry throw after earlier sections had already had their stops replaced.
        var resolved = WordTabStopHelper.Resolve(p.TabStops);

        foreach (var section in sections)
        {
            var footer = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);
            if (resolved.Count > 0) ApplyTabStops(doc, footer, resolved);
        }

        MarkModified(context);
        return new SuccessResult { Message = "Footer tab stops set" };
    }

    /// <summary>
    ///     Applies already-resolved tab stops to the footer paragraph.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="footer">The footer.</param>
    /// <param name="resolved">
    ///     The stops to apply, read by <see cref="WordTabStopHelper" /> before this was called.
    ///     Reading each entry as it was applied meant a malformed one threw after Clear(), leaving
    ///     the paragraph with no stops at all — the defect fixed for `edit` in §23.4 and left
    ///     standing here (R13-W01).
    /// </param>
    private static void ApplyTabStops(
        Document doc, Aspose.Words.HeaderFooter footer, List<TabStop> resolved)
    {
        var para = footer.FirstParagraph ?? new WordParagraph(doc);
        para.ParagraphFormat.TabStops.Clear();
        foreach (var stop in resolved) para.ParagraphFormat.TabStops.Add(stop);

        if (footer.FirstParagraph == null) footer.AppendChild(para);
    }

    /// <summary>
    ///     Extracts parameters for the set footer tabs operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetFooterTabsParameters ExtractSetFooterTabsParameters(OperationParameters parameters)
    {
        return new SetFooterTabsParameters(
            parameters.GetOptional<JsonArray?>("tabStops"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional("headerFooterType", "primary")
        );
    }

    /// <summary>
    ///     Parameters for the set footer tabs operation.
    /// </summary>
    /// <param name="TabStops">The tab stops array.</param>
    /// <param name="SectionIndex">The section index.</param>
    /// <param name="HeaderFooterType">The header/footer type.</param>
    private sealed record SetFooterTabsParameters(
        JsonArray? TabStops,
        int SectionIndex,
        string HeaderFooterType);
}
