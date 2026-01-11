using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

public class SetHeaderFooterHandler : OperationHandlerBase<Document>
{
    public override string Operation => "set_header_footer";

    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var headerLeft = parameters.GetOptional<string?>("headerLeft");
        var headerCenter = parameters.GetOptional<string?>("headerCenter");
        var headerRight = parameters.GetOptional<string?>("headerRight");
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

        var hfHeaderType = WordHeaderFooterHelper.GetHeaderFooterType(headerFooterType, true);
        var hfFooterType = WordHeaderFooterHelper.GetHeaderFooterType(headerFooterType, false);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        var hasHeaderContent = !string.IsNullOrEmpty(headerLeft) || !string.IsNullOrEmpty(headerCenter) ||
                               !string.IsNullOrEmpty(headerRight);
        var hasFooterContent = !string.IsNullOrEmpty(footerLeft) || !string.IsNullOrEmpty(footerCenter) ||
                               !string.IsNullOrEmpty(footerRight);

        foreach (var section in sections)
        {
            if (hasHeaderContent)
            {
                var header = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfHeaderType);

                if (clearExisting)
                {
                    if (clearTextOnly)
                        WordHeaderFooterHelper.ClearTextOnly(header);
                    else
                        header.RemoveAllChildren();
                }

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

            if (hasFooterContent)
            {
                var footer = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfFooterType);

                if (clearExisting)
                {
                    if (clearTextOnly)
                        WordHeaderFooterHelper.ClearTextOnly(footer);
                    else
                        footer.RemoveAllChildren();
                }

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

        return Success("Header and footer set");
    }
}
