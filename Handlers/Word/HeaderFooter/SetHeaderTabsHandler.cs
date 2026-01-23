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

        foreach (var section in sections)
        {
            var header = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);
            if (p.TabStops is { Count: > 0 })
                ApplyTabStops(doc, header, p.TabStops);
        }

        MarkModified(context);
        return new SuccessResult { Message = "Header tab stops set" };
    }

    /// <summary>
    ///     Applies tab stops to the header paragraph.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="header">The header.</param>
    /// <param name="tabStops">The tab stops array.</param>
    private static void ApplyTabStops(Document doc, Aspose.Words.HeaderFooter header, JsonArray tabStops)
    {
        var para = header.FirstParagraph ?? new WordParagraph(doc);
        para.ParagraphFormat.TabStops.Clear();

        foreach (var tabStopJson in tabStops)
        {
            var tabStop = ParseTabStop(tabStopJson);
            para.ParagraphFormat.TabStops.Add(tabStop);
        }

        if (header.FirstParagraph == null) header.AppendChild(para);
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
