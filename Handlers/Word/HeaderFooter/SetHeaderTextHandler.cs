using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

public class SetHeaderTextHandler : OperationHandlerBase<Document>
{
    public override string Operation => "set_header_text";

    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var headerLeft = parameters.GetOptional<string?>("headerLeft");
        var headerCenter = parameters.GetOptional<string?>("headerCenter");
        var headerRight = parameters.GetOptional<string?>("headerRight");
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontNameAscii = parameters.GetOptional<string?>("fontNameAscii");
        var fontNameFarEast = parameters.GetOptional<string?>("fontNameFarEast");
        var fontSize = parameters.GetOptional<double?>("fontSize");
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var headerFooterType = parameters.GetOptional("headerFooterType", "primary");
        var autoTabStops = parameters.GetOptional("autoTabStops", true);
        var clearExisting = parameters.GetOptional("clearExisting", true);
        var clearTextOnly = parameters.GetOptional("clearTextOnly", false);

        var doc = context.Document;

        var hasContent = !string.IsNullOrEmpty(headerLeft) || !string.IsNullOrEmpty(headerCenter) ||
                         !string.IsNullOrEmpty(headerRight);
        if (!hasContent)
            return "Warning: No header text content provided";

        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(headerFooterType, true);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            var header = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

            if (clearExisting)
            {
                if (clearTextOnly)
                    WordHeaderFooterHelper.ClearTextOnly(header);
                else
                    header.RemoveAllChildren();
            }

            if (hasContent)
            {
                if (!clearTextOnly)
                    header.RemoveAllChildren();

                var headerPara = new WordParagraph(doc);
                header.AppendChild(headerPara);

                if (autoTabStops && (!string.IsNullOrEmpty(headerCenter) || !string.IsNullOrEmpty(headerRight)))
                {
                    var pageWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin -
                                    section.PageSetup.RightMargin;
                    headerPara.ParagraphFormat.TabStops.Clear();
                    headerPara.ParagraphFormat.TabStops.Add(new TabStop(pageWidth / 2, TabAlignment.Center,
                        TabLeader.None));
                    headerPara.ParagraphFormat.TabStops.Add(new TabStop(pageWidth, TabAlignment.Right,
                        TabLeader.None));
                }

                var builder = new DocumentBuilder(doc);
                builder.MoveTo(headerPara);

                if (!string.IsNullOrEmpty(headerLeft))
                    WordHeaderFooterHelper.InsertTextOrField(builder, headerLeft, fontName, fontNameAscii,
                        fontNameFarEast, fontSize);

                if (!string.IsNullOrEmpty(headerCenter))
                {
                    builder.Write("\t");
                    WordHeaderFooterHelper.InsertTextOrField(builder, headerCenter, fontName, fontNameAscii,
                        fontNameFarEast, fontSize);
                }

                if (!string.IsNullOrEmpty(headerRight))
                {
                    builder.Write("\t");
                    WordHeaderFooterHelper.InsertTextOrField(builder, headerRight, fontName, fontNameAscii,
                        fontNameFarEast, fontSize);
                }
            }
        }

        MarkModified(context);

        List<string> contentParts = [];
        if (!string.IsNullOrEmpty(headerLeft)) contentParts.Add("left");
        if (!string.IsNullOrEmpty(headerCenter)) contentParts.Add("center");
        if (!string.IsNullOrEmpty(headerRight)) contentParts.Add("right");

        var contentDesc = string.Join(", ", contentParts);
        var sectionsDesc = sectionIndex == -1 ? "all sections" : $"section {sectionIndex}";

        return Success($"Header text set successfully ({contentDesc}) in {sectionsDesc}");
    }
}
