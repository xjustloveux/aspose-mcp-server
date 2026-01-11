using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

public class SetFooterTextHandler : OperationHandlerBase<Document>
{
    public override string Operation => "set_footer_text";

    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var footerLeft = parameters.GetOptional<string?>("footerLeft");
        var footerCenter = parameters.GetOptional<string?>("footerCenter");
        var footerRight = parameters.GetOptional<string?>("footerRight");
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

        var hasContent = !string.IsNullOrEmpty(footerLeft) || !string.IsNullOrEmpty(footerCenter) ||
                         !string.IsNullOrEmpty(footerRight);
        if (!hasContent)
            return "Warning: No footer text content provided";

        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(headerFooterType, false);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        foreach (var section in sections)
        {
            var footer = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

            if (clearExisting)
            {
                if (clearTextOnly)
                    WordHeaderFooterHelper.ClearTextOnly(footer);
                else
                    footer.RemoveAllChildren();
            }

            if (hasContent)
            {
                if (!clearTextOnly)
                    footer.RemoveAllChildren();

                var footerPara = new WordParagraph(doc);
                footer.AppendChild(footerPara);

                if (autoTabStops && (!string.IsNullOrEmpty(footerCenter) || !string.IsNullOrEmpty(footerRight)))
                {
                    var pageWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin -
                                    section.PageSetup.RightMargin;
                    footerPara.ParagraphFormat.TabStops.Clear();
                    footerPara.ParagraphFormat.TabStops.Add(new TabStop(pageWidth / 2, TabAlignment.Center,
                        TabLeader.None));
                    footerPara.ParagraphFormat.TabStops.Add(new TabStop(pageWidth, TabAlignment.Right,
                        TabLeader.None));
                }

                var builder = new DocumentBuilder(doc);
                builder.MoveTo(footerPara);

                if (!string.IsNullOrEmpty(footerLeft))
                    WordHeaderFooterHelper.InsertTextOrField(builder, footerLeft, fontName, fontNameAscii,
                        fontNameFarEast, fontSize);

                if (!string.IsNullOrEmpty(footerCenter))
                {
                    builder.Write("\t");
                    WordHeaderFooterHelper.InsertTextOrField(builder, footerCenter, fontName, fontNameAscii,
                        fontNameFarEast, fontSize);
                }

                if (!string.IsNullOrEmpty(footerRight))
                {
                    builder.Write("\t");
                    WordHeaderFooterHelper.InsertTextOrField(builder, footerRight, fontName, fontNameAscii,
                        fontNameFarEast, fontSize);
                }
            }
        }

        MarkModified(context);

        List<string> contentParts = [];
        if (!string.IsNullOrEmpty(footerLeft)) contentParts.Add("left");
        if (!string.IsNullOrEmpty(footerCenter)) contentParts.Add("center");
        if (!string.IsNullOrEmpty(footerRight)) contentParts.Add("right");

        var contentDesc = string.Join(", ", contentParts);
        var sectionsDesc = sectionIndex == -1 ? "all sections" : $"section {sectionIndex}";

        return Success($"Footer text set successfully ({contentDesc}) in {sectionsDesc}");
    }
}
