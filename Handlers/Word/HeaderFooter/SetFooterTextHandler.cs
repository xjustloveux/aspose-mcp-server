using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting footer text in Word documents.
/// </summary>
public class SetFooterTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_footer_text";

    /// <summary>
    ///     Sets text content in the document footer.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: footerLeft, footerCenter, footerRight, fontName, fontNameAscii, fontNameFarEast,
    ///     fontSize, sectionIndex, headerFooterType, autoTabStops, clearExisting, clearTextOnly
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetFooterTextParameters(parameters);

        var doc = context.Document;

        var hasContent = !string.IsNullOrEmpty(p.FooterLeft) || !string.IsNullOrEmpty(p.FooterCenter) ||
                         !string.IsNullOrEmpty(p.FooterRight);
        if (!hasContent)
            return "Warning: No footer text content provided";

        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType, false);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        foreach (var section in sections)
        {
            var footer = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

            if (p.ClearExisting)
            {
                if (p.ClearTextOnly)
                    WordHeaderFooterHelper.ClearTextOnly(footer);
                else
                    footer.RemoveAllChildren();
            }

            if (hasContent)
            {
                if (!p.ClearTextOnly)
                    footer.RemoveAllChildren();

                var footerPara = new WordParagraph(doc);
                footer.AppendChild(footerPara);

                if (p.AutoTabStops && (!string.IsNullOrEmpty(p.FooterCenter) || !string.IsNullOrEmpty(p.FooterRight)))
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

                if (!string.IsNullOrEmpty(p.FooterLeft))
                    WordHeaderFooterHelper.InsertTextOrField(builder, p.FooterLeft, p.FontName, p.FontNameAscii,
                        p.FontNameFarEast, p.FontSize);

                if (!string.IsNullOrEmpty(p.FooterCenter))
                {
                    builder.Write("\t");
                    WordHeaderFooterHelper.InsertTextOrField(builder, p.FooterCenter, p.FontName, p.FontNameAscii,
                        p.FontNameFarEast, p.FontSize);
                }

                if (!string.IsNullOrEmpty(p.FooterRight))
                {
                    builder.Write("\t");
                    WordHeaderFooterHelper.InsertTextOrField(builder, p.FooterRight, p.FontName, p.FontNameAscii,
                        p.FontNameFarEast, p.FontSize);
                }
            }
        }

        MarkModified(context);

        List<string> contentParts = [];
        if (!string.IsNullOrEmpty(p.FooterLeft)) contentParts.Add("left");
        if (!string.IsNullOrEmpty(p.FooterCenter)) contentParts.Add("center");
        if (!string.IsNullOrEmpty(p.FooterRight)) contentParts.Add("right");

        var contentDesc = string.Join(", ", contentParts);
        var sectionsDesc = p.SectionIndex == -1 ? "all sections" : $"section {p.SectionIndex}";

        return Success($"Footer text set successfully ({contentDesc}) in {sectionsDesc}");
    }

    /// <summary>
    ///     Extracts parameters for the set footer text operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetFooterTextParameters ExtractSetFooterTextParameters(OperationParameters parameters)
    {
        return new SetFooterTextParameters(
            parameters.GetOptional<string?>("footerLeft"),
            parameters.GetOptional<string?>("footerCenter"),
            parameters.GetOptional<string?>("footerRight"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<string?>("fontNameAscii"),
            parameters.GetOptional<string?>("fontNameFarEast"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional("headerFooterType", "primary"),
            parameters.GetOptional("autoTabStops", true),
            parameters.GetOptional("clearExisting", true),
            parameters.GetOptional("clearTextOnly", false)
        );
    }

    /// <summary>
    ///     Parameters for the set footer text operation.
    /// </summary>
    /// <param name="FooterLeft">The left footer text.</param>
    /// <param name="FooterCenter">The center footer text.</param>
    /// <param name="FooterRight">The right footer text.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontNameAscii">The ASCII font name.</param>
    /// <param name="FontNameFarEast">The Far East font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="SectionIndex">The section index.</param>
    /// <param name="HeaderFooterType">The header/footer type.</param>
    /// <param name="AutoTabStops">Whether to auto-create tab stops.</param>
    /// <param name="ClearExisting">Whether to clear existing content.</param>
    /// <param name="ClearTextOnly">Whether to clear text only.</param>
    private record SetFooterTextParameters(
        string? FooterLeft,
        string? FooterCenter,
        string? FooterRight,
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        int SectionIndex,
        string HeaderFooterType,
        bool AutoTabStops,
        bool ClearExisting,
        bool ClearTextOnly);
}
