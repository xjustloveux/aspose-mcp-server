using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

/// <summary>
///     Handler for setting both headers and footers in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetHeaderFooterHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_header_footer";

    /// <summary>
    ///     Sets both header and footer content in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: headerLeft, headerCenter, headerRight, footerLeft, footerCenter, footerRight,
    ///     fontName, fontNameAscii, fontNameFarEast, fontSize, sectionIndex, headerFooterType,
    ///     autoTabStops, clearExisting, clearTextOnly
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetHeaderFooterParameters(parameters);

        var doc = context.Document;
        var fontSettings = new FontSettings(p.FontName, p.FontNameAscii, p.FontNameFarEast, p.FontSize);

        var hfHeaderType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType ?? "primary", true);
        var hfFooterType = WordHeaderFooterHelper.GetHeaderFooterType(p.HeaderFooterType ?? "primary", false);
        var sections = p.SectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[p.SectionIndex]];

        var hasHeaderContent = HasContent(p.HeaderLeft, p.HeaderCenter, p.HeaderRight);
        var hasFooterContent = HasContent(p.FooterLeft, p.FooterCenter, p.FooterRight);

        var headerContent = new HeaderFooterContent(p.HeaderLeft, p.HeaderCenter, p.HeaderRight);
        var footerContent = new HeaderFooterContent(p.FooterLeft, p.FooterCenter, p.FooterRight);
        var settings = new HeaderFooterSettings(fontSettings, p.AutoTabStops, p.ClearExisting, p.ClearTextOnly);

        foreach (var section in sections)
        {
            if (hasHeaderContent)
                SetupHeaderFooterContent(section, doc, hfHeaderType, headerContent, settings);

            if (hasFooterContent)
                SetupHeaderFooterContent(section, doc, hfFooterType, footerContent, settings);
        }

        MarkModified(context);

        return new SuccessResult { Message = "Header and footer set" };
    }

    /// <summary>
    ///     Checks if any content is provided.
    /// </summary>
    /// <param name="left">The left content.</param>
    /// <param name="center">The center content.</param>
    /// <param name="right">The right content.</param>
    /// <returns>True if any content is provided.</returns>
    private static bool HasContent(string? left, string? center, string? right)
    {
        return !string.IsNullOrEmpty(left) || !string.IsNullOrEmpty(center) || !string.IsNullOrEmpty(right);
    }

    /// <summary>
    ///     Sets up header/footer content in a section.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <param name="doc">The Word document.</param>
    /// <param name="hfType">The header/footer type.</param>
    /// <param name="content">The header/footer content (left, center, right).</param>
    /// <param name="settings">The header/footer settings.</param>
    private static void SetupHeaderFooterContent(Section section, Document doc, HeaderFooterType hfType,
        HeaderFooterContent content, HeaderFooterSettings settings)
    {
        var hf = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

        ClearHeaderFooter(hf, settings.ClearExisting, settings.ClearTextOnly);

        var para = new WordParagraph(doc);
        hf.AppendChild(para);

        if (settings.AutoTabStops && (!string.IsNullOrEmpty(content.Center) || !string.IsNullOrEmpty(content.Right)))
            SetupTabStops(section, para);

        InsertContent(doc, para, content, settings.FontSettings);
    }

    /// <summary>
    ///     Clears header/footer content based on settings.
    /// </summary>
    /// <param name="hf">The header/footer.</param>
    /// <param name="clearExisting">Whether to clear existing content.</param>
    /// <param name="clearTextOnly">Whether to clear text only.</param>
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
    ///     Sets up tab stops for center and right alignment.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <param name="para">The paragraph.</param>
    private static void SetupTabStops(Section section, WordParagraph para)
    {
        var pageWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
        para.ParagraphFormat.TabStops.Clear();
        para.ParagraphFormat.TabStops.Add(new TabStop(pageWidth / 2, TabAlignment.Center, TabLeader.None));
        para.ParagraphFormat.TabStops.Add(new TabStop(pageWidth, TabAlignment.Right, TabLeader.None));
    }

    /// <summary>
    ///     Inserts content into the paragraph.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="para">The paragraph.</param>
    /// <param name="content">The header/footer content (left, center, right).</param>
    /// <param name="fontSettings">The font settings.</param>
    private static void InsertContent(Document doc, WordParagraph para, HeaderFooterContent content,
        FontSettings fontSettings)
    {
        var builder = new DocumentBuilder(doc);
        builder.MoveTo(para);

        if (!string.IsNullOrEmpty(content.Left))
            WordHeaderFooterHelper.InsertTextOrField(builder, content.Left, fontSettings);

        if (!string.IsNullOrEmpty(content.Center))
        {
            builder.Write("\t");
            WordHeaderFooterHelper.InsertTextOrField(builder, content.Center, fontSettings);
        }

        if (!string.IsNullOrEmpty(content.Right))
        {
            builder.Write("\t");
            WordHeaderFooterHelper.InsertTextOrField(builder, content.Right, fontSettings);
        }
    }

    /// <summary>
    ///     Extracts parameters for the set header footer operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetHeaderFooterParameters ExtractSetHeaderFooterParameters(OperationParameters parameters)
    {
        return new SetHeaderFooterParameters(
            parameters.GetOptional<string?>("headerLeft"),
            parameters.GetOptional<string?>("headerCenter"),
            parameters.GetOptional<string?>("headerRight"),
            parameters.GetOptional<string?>("footerLeft"),
            parameters.GetOptional<string?>("footerCenter"),
            parameters.GetOptional<string?>("footerRight"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<string?>("fontNameAscii"),
            parameters.GetOptional<string?>("fontNameFarEast"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional("sectionIndex", -1),
            parameters.GetOptional<string?>("headerFooterType"),
            parameters.GetOptional("autoTabStops", true),
            parameters.GetOptional("clearExisting", true),
            parameters.GetOptional("clearTextOnly", false)
        );
    }

    /// <summary>
    ///     Record to hold header/footer content (left, center, right).
    /// </summary>
    /// <param name="Left">The left content.</param>
    /// <param name="Center">The center content.</param>
    /// <param name="Right">The right content.</param>
    private sealed record HeaderFooterContent(string? Left, string? Center, string? Right);

    /// <summary>
    ///     Record to hold header/footer settings.
    /// </summary>
    /// <param name="FontSettings">The font settings.</param>
    /// <param name="AutoTabStops">Whether to auto-create tab stops.</param>
    /// <param name="ClearExisting">Whether to clear existing content.</param>
    /// <param name="ClearTextOnly">Whether to clear text only.</param>
    private sealed record HeaderFooterSettings(
        FontSettings FontSettings,
        bool AutoTabStops,
        bool ClearExisting,
        bool ClearTextOnly);

    /// <summary>
    ///     Parameters for the set header footer operation.
    /// </summary>
    /// <param name="HeaderLeft">The left header text.</param>
    /// <param name="HeaderCenter">The center header text.</param>
    /// <param name="HeaderRight">The right header text.</param>
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
    private sealed record SetHeaderFooterParameters(
        string? HeaderLeft,
        string? HeaderCenter,
        string? HeaderRight,
        string? FooterLeft,
        string? FooterCenter,
        string? FooterRight,
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        int SectionIndex,
        string? HeaderFooterType,
        bool AutoTabStops,
        bool ClearExisting,
        bool ClearTextOnly);
}
