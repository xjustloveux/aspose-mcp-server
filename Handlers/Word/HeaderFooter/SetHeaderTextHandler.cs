using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting header text in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetHeaderTextParameters(parameters);

        var doc = context.Document;

        var hasContent = !string.IsNullOrEmpty(p.HeaderLeft) || !string.IsNullOrEmpty(p.HeaderCenter) ||
                         !string.IsNullOrEmpty(p.HeaderRight);
        if (!hasContent)
            return new SuccessResult { Message = "Warning: No header text content provided" };

        var hfType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType, true);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        foreach (var section in sections)
            ProcessSection(doc, section, hfType, p);

        MarkModified(context);

        return new SuccessResult
            { Message = BuildResultMessage("Header", p.HeaderLeft, p.HeaderCenter, p.HeaderRight, p.SectionIndex) };
    }

    /// <summary>
    ///     Processes a single section to set header text.
    /// </summary>
    private static void ProcessSection(Document doc, Section section, HeaderFooterType hfType,
        SetHeaderTextParameters p)
    {
        var header = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

        ClearHeaderFooter(header, p.ClearExisting, p.ClearTextOnly);

        var headerPara = new WordParagraph(doc);
        header.AppendChild(headerPara);

        ConfigureTabStops(section, headerPara, p.AutoTabStops, p.HeaderCenter, p.HeaderRight);

        var builder = new DocumentBuilder(doc);
        builder.MoveTo(headerPara);

        InsertTextContent(builder, p.HeaderLeft, p.HeaderCenter, p.HeaderRight, p.FontSettings);
    }

    /// <summary>
    ///     Clears header/footer content based on parameters.
    /// </summary>
    private static void ClearHeaderFooter(Aspose.Words.HeaderFooter hf, bool clearExisting, bool clearTextOnly)
    {
        if (clearExisting)
        {
            if (clearTextOnly)
                WordHeaderFooterHelper.ClearTextOnly(hf);
            else
                hf.RemoveAllChildren();
        }

        if (!clearTextOnly)
            hf.RemoveAllChildren();
    }

    /// <summary>
    ///     Configures tab stops for center and right alignment.
    /// </summary>
    private static void ConfigureTabStops(Section section, WordParagraph para, bool autoTabStops,
        string? centerText, string? rightText)
    {
        if (!autoTabStops || (string.IsNullOrEmpty(centerText) && string.IsNullOrEmpty(rightText)))
            return;

        var pageWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
        para.ParagraphFormat.TabStops.Clear();
        para.ParagraphFormat.TabStops.Add(new TabStop(pageWidth / 2, TabAlignment.Center, TabLeader.None));
        para.ParagraphFormat.TabStops.Add(new TabStop(pageWidth, TabAlignment.Right, TabLeader.None));
    }

    /// <summary>
    ///     Inserts text content with optional tab separators.
    /// </summary>
    private static void InsertTextContent(DocumentBuilder builder, string? left, string? center, string? right,
        FontSettings fontSettings)
    {
        if (!string.IsNullOrEmpty(left))
            WordHeaderFooterHelper.InsertTextOrField(builder, left, fontSettings);

        if (!string.IsNullOrEmpty(center))
        {
            builder.Write("\t");
            WordHeaderFooterHelper.InsertTextOrField(builder, center, fontSettings);
        }

        if (!string.IsNullOrEmpty(right))
        {
            builder.Write("\t");
            WordHeaderFooterHelper.InsertTextOrField(builder, right, fontSettings);
        }
    }

    /// <summary>
    ///     Builds the result message.
    /// </summary>
    private static string BuildResultMessage(string type, string? left, string? center, string? right, int sectionIndex)
    {
        List<string> contentParts = [];
        if (!string.IsNullOrEmpty(left)) contentParts.Add("left");
        if (!string.IsNullOrEmpty(center)) contentParts.Add("center");
        if (!string.IsNullOrEmpty(right)) contentParts.Add("right");

        var contentDesc = string.Join(", ", contentParts);
        var sectionsDesc = sectionIndex == -1 ? "all sections" : $"section {sectionIndex}";

        return $"{type} text set successfully ({contentDesc}) in {sectionsDesc}";
    }

    /// <summary>
    ///     Extracts parameters for the set header text operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetHeaderTextParameters ExtractSetHeaderTextParameters(OperationParameters parameters)
    {
        var fontSettings = new FontSettings(
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<string?>("fontNameAscii"),
            parameters.GetOptional<string?>("fontNameFarEast"),
            parameters.GetOptional<double?>("fontSize"));

        return new SetHeaderTextParameters(
            parameters.GetOptional<string?>("headerLeft"),
            parameters.GetOptional<string?>("headerCenter"),
            parameters.GetOptional<string?>("headerRight"),
            fontSettings,
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
    /// <param name="FontSettings">The font settings.</param>
    /// <param name="SectionIndex">The section index.</param>
    /// <param name="HeaderFooterType">The header/footer type.</param>
    /// <param name="AutoTabStops">Whether to auto-create tab stops.</param>
    /// <param name="ClearExisting">Whether to clear existing content.</param>
    /// <param name="ClearTextOnly">Whether to clear text only.</param>
    private sealed record SetHeaderTextParameters(
        string? HeaderLeft,
        string? HeaderCenter,
        string? HeaderRight,
        FontSettings FontSettings,
        int SectionIndex,
        string HeaderFooterType,
        bool AutoTabStops,
        bool ClearExisting,
        bool ClearTextOnly);
}
