using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting footer tab stops in Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetFooterTabsParameters(parameters);

        var doc = context.Document;
        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType, false);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        foreach (var section in sections)
        {
            var footer = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);
            if (p.TabStops is { Count: > 0 })
                ApplyTabStops(doc, footer, p.TabStops);
        }

        MarkModified(context);
        return Success("Footer tab stops set");
    }

    /// <summary>
    ///     Applies tab stops to the footer paragraph.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="footer">The footer.</param>
    /// <param name="tabStops">The tab stops array.</param>
    private static void ApplyTabStops(Document doc, Aspose.Words.HeaderFooter footer, JsonArray tabStops)
    {
        var para = footer.FirstParagraph ?? new WordParagraph(doc);
        para.ParagraphFormat.TabStops.Clear();

        foreach (var tabStopJson in tabStops)
        {
            var tabStop = ParseTabStop(tabStopJson);
            para.ParagraphFormat.TabStops.Add(tabStop);
        }

        if (footer.FirstParagraph == null) footer.AppendChild(para);
    }

    /// <summary>
    ///     Parses a tab stop from JSON.
    /// </summary>
    /// <param name="tabStopJson">The tab stop JSON node.</param>
    /// <returns>The parsed tab stop.</returns>
    private static TabStop ParseTabStop(JsonNode? tabStopJson)
    {
        var position = tabStopJson?["position"]?.GetValue<double>() ?? 0;
        var alignmentStr = tabStopJson?["alignment"]?.GetValue<string>() ?? "left";
        var leaderStr = tabStopJson?["leader"]?.GetValue<string>() ?? "none";

        var tabAlignment = GetTabAlignment(alignmentStr);
        var tabLeader = GetTabLeader(leaderStr);

        return new TabStop(position, tabAlignment, tabLeader);
    }

    /// <summary>
    ///     Gets the tab alignment from alignment string.
    /// </summary>
    /// <param name="alignment">The alignment string.</param>
    /// <returns>The tab alignment.</returns>
    private static TabAlignment GetTabAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "center" => TabAlignment.Center,
            "right" => TabAlignment.Right,
            "decimal" => TabAlignment.Decimal,
            "bar" => TabAlignment.Bar,
            _ => TabAlignment.Left
        };
    }

    /// <summary>
    ///     Gets the tab leader from leader string.
    /// </summary>
    /// <param name="leader">The leader string.</param>
    /// <returns>The tab leader.</returns>
    private static TabLeader GetTabLeader(string leader)
    {
        return leader.ToLower() switch
        {
            "dots" => TabLeader.Dots,
            "dashes" => TabLeader.Dashes,
            "line" => TabLeader.Line,
            _ => TabLeader.None
        };
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
