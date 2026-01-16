using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for getting tab stops in Word documents.
/// </summary>
public class GetTabStopsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_tab_stops";

    /// <summary>
    ///     Gets tab stops for a paragraph.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    ///     Optional: location, sectionIndex, allParagraphs, includeStyle
    /// </param>
    /// <returns>A JSON string containing the tab stops information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var location = parameters.GetOptional("location", "body");
        var paragraphIndex = parameters.GetOptional("paragraphIndex", 0);
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var allParagraphs = parameters.GetOptional("allParagraphs", false);
        var includeStyle = parameters.GetOptional("includeStyle", true);

        var doc = context.Document;

        if (sectionIndex >= doc.Sections.Count)
            throw new ArgumentException(
                $"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");

        var section = doc.Sections[sectionIndex];
        var (targetParagraphs, locationDesc) = GetTargetParagraphs(section, location, paragraphIndex, allParagraphs);

        if (targetParagraphs.Count == 0)
            throw new InvalidOperationException("No target paragraphs found");

        var allTabStops = CollectTabStops(targetParagraphs, allParagraphs, includeStyle);

        var tabStopsList = allTabStops
            .OrderBy(x => x.Value.position)
            .Select(kvp => new
            {
                positionPt = kvp.Value.position,
                positionCm = Math.Round(kvp.Value.position / 28.35, 2),
                alignment = kvp.Value.alignment.ToString(),
                leader = kvp.Value.leader.ToString(),
                kvp.Value.source
            })
            .ToList();

        var result = new
        {
            location,
            locationDescription = locationDesc,
            sectionIndex,
            paragraphIndex = location == "body" && !allParagraphs ? paragraphIndex : (int?)null,
            allParagraphs,
            includeStyle,
            paragraphCount = targetParagraphs.Count,
            count = allTabStops.Count,
            tabStops = tabStopsList
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }

    private static (List<WordParagraph> paragraphs, string locationDesc) GetTargetParagraphs(
        Section section, string location, int paragraphIndex, bool allParagraphs)
    {
        return location.ToLower() switch
        {
            "header" => GetHeaderParagraphs(section, allParagraphs),
            "footer" => GetFooterParagraphs(section, allParagraphs),
            _ => GetBodyParagraphs(section, paragraphIndex, allParagraphs)
        };
    }

    private static (List<WordParagraph>, string) GetHeaderParagraphs(Section section, bool allParagraphs)
    {
        var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
            throw new InvalidOperationException("Header not found");

        var headerParas = header.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var paragraphs = allParagraphs ? headerParas : headerParas.Count > 0 ? [headerParas[0]] : [];
        return (paragraphs, "Header");
    }

    private static (List<WordParagraph>, string) GetFooterParagraphs(Section section, bool allParagraphs)
    {
        var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
        if (footer == null)
            throw new InvalidOperationException("Footer not found");

        var footerParas = footer.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var paragraphs = allParagraphs ? footerParas : footerParas.Count > 0 ? [footerParas[0]] : [];
        return (paragraphs, "Footer");
    }

    private static (List<WordParagraph>, string) GetBodyParagraphs(Section section, int paragraphIndex,
        bool allParagraphs)
    {
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        if (allParagraphs)
            return (paragraphs, "Body");

        if (paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} out of range (total paragraphs: {paragraphs.Count})");

        return ([paragraphs[paragraphIndex]], $"Body Paragraph {paragraphIndex}");
    }

    private static Dictionary<string, (double position, TabAlignment alignment, TabLeader leader, string source)>
        CollectTabStops(List<WordParagraph> targetParagraphs, bool allParagraphs, bool includeStyle)
    {
        var allTabStops =
            new Dictionary<string, (double position, TabAlignment alignment, TabLeader leader, string source)>();

        for (var paraIdx = 0; paraIdx < targetParagraphs.Count; paraIdx++)
        {
            var para = targetParagraphs[paraIdx];
            var paraSource = allParagraphs ? $"Paragraph {paraIdx}" : "Paragraph";

            CollectParagraphTabStops(para, paraSource, allTabStops);

            if (includeStyle && para.ParagraphFormat.Style != null)
                CollectStyleTabStops(para, paraSource, allTabStops);
        }

        return allTabStops;
    }

    private static void CollectParagraphTabStops(WordParagraph para, string paraSource,
        Dictionary<string, (double, TabAlignment, TabLeader, string)> allTabStops)
    {
        var paraTabStops = para.ParagraphFormat.TabStops;
        for (var i = 0; i < paraTabStops.Count; i++)
        {
            var tab = paraTabStops[i];
            var position = Math.Round(tab.Position, 2);
            var key = $"{position}_{tab.Alignment}";
            if (!allTabStops.ContainsKey(key))
                allTabStops[key] = (position, tab.Alignment, tab.Leader, $"{paraSource} (Custom)");
        }
    }

    private static void CollectStyleTabStops(WordParagraph para, string paraSource,
        Dictionary<string, (double, TabAlignment, TabLeader, string)> allTabStops)
    {
        var styleChain = BuildStyleChain(para);

        foreach (var chainStyle in styleChain)
        {
            if (chainStyle.ParagraphFormat == null) continue;

            var styleTabStops = chainStyle.ParagraphFormat.TabStops;
            for (var i = 0; i < styleTabStops.Count; i++)
            {
                var tab = styleTabStops[i];
                var position = Math.Round(tab.Position, 2);
                var key = $"{position}_{tab.Alignment}";

                if (allTabStops.ContainsKey(key)) continue;

                var styleName = chainStyle == para.ParagraphFormat.Style
                    ? chainStyle.Name
                    : $"{para.ParagraphFormat.Style.Name} (Base: {chainStyle.Name})";
                allTabStops[key] = (position, tab.Alignment, tab.Leader, $"{paraSource} (Style: {styleName})");
            }
        }
    }

    private static List<Style> BuildStyleChain(WordParagraph para)
    {
        List<Style> styleChain = [];
        var currentStyle = para.ParagraphFormat.Style;

        while (currentStyle != null)
        {
            styleChain.Add(currentStyle);
            currentStyle = GetBaseStyle(para.Document, currentStyle, styleChain);
        }

        return styleChain;
    }

    private static Style? GetBaseStyle(DocumentBase doc, Style currentStyle, List<Style> styleChain)
    {
        if (string.IsNullOrEmpty(currentStyle.BaseStyleName))
            return null;

        try
        {
            var baseStyle = doc.Styles[currentStyle.BaseStyleName];
            if (baseStyle != null && !styleChain.Contains(baseStyle))
                return baseStyle;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WARN] Error accessing paragraph style: {ex.Message}");
        }

        return null;
    }
}
