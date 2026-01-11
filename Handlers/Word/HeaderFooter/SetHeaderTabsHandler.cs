using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

public class SetHeaderTabsHandler : OperationHandlerBase<Document>
{
    public override string Operation => "set_header_tabs";

    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var tabStops = parameters.GetOptional<JsonArray?>("tabStops");
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var headerFooterType = parameters.GetOptional("headerFooterType", "primary");

        var doc = context.Document;
        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(headerFooterType, true);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            var header = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

            if (tabStops is { Count: > 0 })
            {
                var para = header.FirstParagraph ?? new WordParagraph(doc);
                para.ParagraphFormat.TabStops.Clear();

                foreach (var tabStopJson in tabStops)
                {
                    var position = tabStopJson?["position"]?.GetValue<double>() ?? 0;
                    var alignmentStr = tabStopJson?["alignment"]?.GetValue<string>() ?? "left";
                    var leaderStr = tabStopJson?["leader"]?.GetValue<string>() ?? "none";

                    var tabAlignment = alignmentStr.ToLower() switch
                    {
                        "center" => TabAlignment.Center,
                        "right" => TabAlignment.Right,
                        "decimal" => TabAlignment.Decimal,
                        "bar" => TabAlignment.Bar,
                        _ => TabAlignment.Left
                    };

                    var tabLeader = leaderStr.ToLower() switch
                    {
                        "dots" => TabLeader.Dots,
                        "dashes" => TabLeader.Dashes,
                        "line" => TabLeader.Line,
                        _ => TabLeader.None
                    };

                    para.ParagraphFormat.TabStops.Add(new TabStop(position, tabAlignment, tabLeader));
                }

                if (header.FirstParagraph == null) header.AppendChild(para);
            }
        }

        MarkModified(context);

        return Success("Header tab stops set");
    }
}
