using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting header text in Word documents.
/// </summary>
public class SetHeaderTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_header_text";

    /// <summary>
    ///     Sets text content in the document header.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: headerLeft, headerCenter, headerRight, fontName, fontNameAscii, fontNameFarEast,
    ///     fontSize, sectionIndex, headerFooterType, autoTabStops, clearExisting, clearTextOnly
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetHeaderTextParameters(parameters);

        var doc = context.Document;

        var hasContent = !string.IsNullOrEmpty(p.HeaderLeft) || !string.IsNullOrEmpty(p.HeaderCenter) ||
                         !string.IsNullOrEmpty(p.HeaderRight);
        if (!hasContent)
            return "Warning: No header text content provided";

        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType, true);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        foreach (var section in sections)
        {
            var header = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

            if (p.ClearExisting)
            {
                if (p.ClearTextOnly)
                    WordHeaderFooterHelper.ClearTextOnly(header);
                else
                    header.RemoveAllChildren();
            }

            if (hasContent)
            {
                if (!p.ClearTextOnly)
                    header.RemoveAllChildren();

                var headerPara = new WordParagraph(doc);
                header.AppendChild(headerPara);

                if (p.AutoTabStops && (!string.IsNullOrEmpty(p.HeaderCenter) || !string.IsNullOrEmpty(p.HeaderRight)))
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

                if (!string.IsNullOrEmpty(p.HeaderLeft))
                    WordHeaderFooterHelper.InsertTextOrField(builder, p.HeaderLeft, p.FontName, p.FontNameAscii,
                        p.FontNameFarEast, p.FontSize);

                if (!string.IsNullOrEmpty(p.HeaderCenter))
                {
                    builder.Write("\t");
                    WordHeaderFooterHelper.InsertTextOrField(builder, p.HeaderCenter, p.FontName, p.FontNameAscii,
                        p.FontNameFarEast, p.FontSize);
                }

                if (!string.IsNullOrEmpty(p.HeaderRight))
                {
                    builder.Write("\t");
                    WordHeaderFooterHelper.InsertTextOrField(builder, p.HeaderRight, p.FontName, p.FontNameAscii,
                        p.FontNameFarEast, p.FontSize);
                }
            }
        }

        MarkModified(context);

        List<string> contentParts = [];
        if (!string.IsNullOrEmpty(p.HeaderLeft)) contentParts.Add("left");
        if (!string.IsNullOrEmpty(p.HeaderCenter)) contentParts.Add("center");
        if (!string.IsNullOrEmpty(p.HeaderRight)) contentParts.Add("right");

        var contentDesc = string.Join(", ", contentParts);
        var sectionsDesc = p.SectionIndex == -1 ? "all sections" : $"section {p.SectionIndex}";

        return Success($"Header text set successfully ({contentDesc}) in {sectionsDesc}");
    }

    /// <summary>
    ///     Extracts parameters for the set header text operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetHeaderTextParameters ExtractSetHeaderTextParameters(OperationParameters parameters)
    {
        return new SetHeaderTextParameters(
            parameters.GetOptional<string?>("headerLeft"),
            parameters.GetOptional<string?>("headerCenter"),
            parameters.GetOptional<string?>("headerRight"),
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
    ///     Parameters for the set header text operation.
    /// </summary>
    /// <param name="HeaderLeft">The left header text.</param>
    /// <param name="HeaderCenter">The center header text.</param>
    /// <param name="HeaderRight">The right header text.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontNameAscii">The ASCII font name.</param>
    /// <param name="FontNameFarEast">The Far East font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="SectionIndex">The section index.</param>
    /// <param name="HeaderFooterType">The header/footer type.</param>
    /// <param name="AutoTabStops">Whether to auto-create tab stops.</param>
    /// <param name="ClearExisting">Whether to clear existing content.</param>
    /// <param name="ClearTextOnly">Whether to clear text only.</param>
    private record SetHeaderTextParameters(
        string? HeaderLeft,
        string? HeaderCenter,
        string? HeaderRight,
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
