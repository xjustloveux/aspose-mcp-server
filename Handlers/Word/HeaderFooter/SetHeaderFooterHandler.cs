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
        var fontSettings = new FontSettings(fontName, fontNameAscii, fontNameFarEast, fontSize);

        var hfHeaderType = WordHeaderFooterHelper.GetHeaderFooterType(headerFooterType, true);
        var hfFooterType = WordHeaderFooterHelper.GetHeaderFooterType(headerFooterType, false);
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : [doc.Sections[sectionIndex]];

        var hasHeaderContent = HasContent(headerLeft, headerCenter, headerRight);
        var hasFooterContent = HasContent(footerLeft, footerCenter, footerRight);

        foreach (var section in sections)
        {
            if (hasHeaderContent)
                SetupHeaderFooterContent(section, doc, hfHeaderType, headerLeft, headerCenter, headerRight,
                    fontSettings, autoTabStops, clearExisting, clearTextOnly);

            if (hasFooterContent)
                SetupHeaderFooterContent(section, doc, hfFooterType, footerLeft, footerCenter, footerRight,
                    fontSettings, autoTabStops, clearExisting, clearTextOnly);
        }

        MarkModified(context);

        return Success("Header and footer set");
    }

    private static bool HasContent(string? left, string? center, string? right)
    {
        return !string.IsNullOrEmpty(left) || !string.IsNullOrEmpty(center) || !string.IsNullOrEmpty(right);
    }

    private static void SetupHeaderFooterContent(Section section, Document doc, HeaderFooterType hfType,
        string? left, string? center, string? right, FontSettings fontSettings,
        bool autoTabStops, bool clearExisting, bool clearTextOnly)
    {
        var hf = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, hfType);

        ClearHeaderFooter(hf, clearExisting, clearTextOnly);

        var para = new WordParagraph(doc);
        hf.AppendChild(para);

        if (autoTabStops && (!string.IsNullOrEmpty(center) || !string.IsNullOrEmpty(right)))
            SetupTabStops(section, para);

        InsertContent(doc, para, left, center, right, fontSettings);
    }

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

    private static void SetupTabStops(Section section, WordParagraph para)
    {
        var pageWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
        para.ParagraphFormat.TabStops.Clear();
        para.ParagraphFormat.TabStops.Add(new TabStop(pageWidth / 2, TabAlignment.Center, TabLeader.None));
        para.ParagraphFormat.TabStops.Add(new TabStop(pageWidth, TabAlignment.Right, TabLeader.None));
    }

    private static void InsertContent(Document doc, WordParagraph para, string? left, string? center, string? right,
        FontSettings fontSettings)
    {
        var builder = new DocumentBuilder(doc);
        builder.MoveTo(para);

        if (!string.IsNullOrEmpty(left))
            WordHeaderFooterHelper.InsertTextOrField(builder, left, fontSettings.FontName, fontSettings.FontNameAscii,
                fontSettings.FontNameFarEast, fontSettings.FontSize);

        if (!string.IsNullOrEmpty(center))
        {
            builder.Write("\t");
            WordHeaderFooterHelper.InsertTextOrField(builder, center, fontSettings.FontName, fontSettings.FontNameAscii,
                fontSettings.FontNameFarEast, fontSettings.FontSize);
        }

        if (!string.IsNullOrEmpty(right))
        {
            builder.Write("\t");
            WordHeaderFooterHelper.InsertTextOrField(builder, right, fontSettings.FontName, fontSettings.FontNameAscii,
                fontSettings.FontNameFarEast, fontSettings.FontSize);
        }
    }

    private record FontSettings(string? FontName, string? FontNameAscii, string? FontNameFarEast, double? FontSize);
}
