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
        var p = ExtractGetTabStopsParameters(parameters);

        var doc = context.Document;

        if (p.SectionIndex >= doc.Sections.Count)
            throw new ArgumentException(
                $"Section index {p.SectionIndex} out of range (total sections: {doc.Sections.Count})");

        var section = doc.Sections[p.SectionIndex];
        var (targetParagraphs, locationDesc) =
            GetTargetParagraphs(section, p.Location, p.ParagraphIndex, p.AllParagraphs);

        if (targetParagraphs.Count == 0)
            throw new InvalidOperationException("No target paragraphs found");

        var allTabStops = CollectTabStops(targetParagraphs, p.AllParagraphs, p.IncludeStyle);

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
            location = p.Location,
            locationDescription = locationDesc,
            sectionIndex = p.SectionIndex,
            paragraphIndex = p is { Location: "body", AllParagraphs: false } ? p.ParagraphIndex : (int?)null,
            allParagraphs = p.AllParagraphs,
            includeStyle = p.IncludeStyle,
            paragraphCount = targetParagraphs.Count,
            count = allTabStops.Count,
            tabStops = tabStopsList
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }

    /// <summary>
    ///     Gets the target paragraphs based on location and settings.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <param name="location">The location type (body, header, footer).</param>
    /// <param name="paragraphIndex">The paragraph index for body location.</param>
    /// <param name="allParagraphs">Whether to get all paragraphs.</param>
    /// <returns>A tuple containing the list of paragraphs and location description.</returns>
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

    /// <summary>
    ///     Gets paragraphs from the header.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <param name="allParagraphs">Whether to get all paragraphs.</param>
    /// <returns>A tuple containing the list of paragraphs and location description.</returns>
    /// <exception cref="InvalidOperationException">Thrown when header is not found.</exception>
    private static (List<WordParagraph>, string) GetHeaderParagraphs(Section section, bool allParagraphs)
    {
        var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
            throw new InvalidOperationException("Header not found");

        var headerParas = header.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var paragraphs = allParagraphs ? headerParas : headerParas.Count > 0 ? [headerParas[0]] : [];
        return (paragraphs, "Header");
    }

    /// <summary>
    ///     Gets paragraphs from the footer.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <param name="allParagraphs">Whether to get all paragraphs.</param>
    /// <returns>A tuple containing the list of paragraphs and location description.</returns>
    /// <exception cref="InvalidOperationException">Thrown when footer is not found.</exception>
    private static (List<WordParagraph>, string) GetFooterParagraphs(Section section, bool allParagraphs)
    {
        var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
        if (footer == null)
            throw new InvalidOperationException("Footer not found");

        var footerParas = footer.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var paragraphs = allParagraphs ? footerParas : footerParas.Count > 0 ? [footerParas[0]] : [];
        return (paragraphs, "Footer");
    }

    /// <summary>
    ///     Gets paragraphs from the body.
    /// </summary>
    /// <param name="section">The document section.</param>
    /// <param name="paragraphIndex">The paragraph index.</param>
    /// <param name="allParagraphs">Whether to get all paragraphs.</param>
    /// <returns>A tuple containing the list of paragraphs and location description.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
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

    /// <summary>
    ///     Collects all tab stops from the target paragraphs.
    /// </summary>
    /// <param name="targetParagraphs">The list of paragraphs to collect from.</param>
    /// <param name="allParagraphs">Whether all paragraphs are included.</param>
    /// <param name="includeStyle">Whether to include style tab stops.</param>
    /// <returns>A dictionary of tab stops keyed by position and alignment.</returns>
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

    /// <summary>
    ///     Collects tab stops defined directly on the paragraph.
    /// </summary>
    /// <param name="para">The paragraph to collect from.</param>
    /// <param name="paraSource">The paragraph source description.</param>
    /// <param name="allTabStops">The dictionary to add tab stops to.</param>
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

    /// <summary>
    ///     Collects tab stops defined in the paragraph's style chain.
    /// </summary>
    /// <param name="para">The paragraph to collect from.</param>
    /// <param name="paraSource">The paragraph source description.</param>
    /// <param name="allTabStops">The dictionary to add tab stops to.</param>
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

    /// <summary>
    ///     Builds the style inheritance chain for a paragraph.
    /// </summary>
    /// <param name="para">The paragraph to build the chain for.</param>
    /// <returns>A list of styles in the inheritance chain.</returns>
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

    /// <summary>
    ///     Gets the base style of the current style if it exists.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="currentStyle">The current style.</param>
    /// <param name="styleChain">The existing style chain to check for cycles.</param>
    /// <returns>The base style or null if none exists.</returns>
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

    /// <summary>
    ///     Extracts parameters for the get tab stops operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetTabStopsParameters ExtractGetTabStopsParameters(OperationParameters parameters)
    {
        return new GetTabStopsParameters(
            parameters.GetOptional("location", "body"),
            parameters.GetOptional("paragraphIndex", 0),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional("allParagraphs", false),
            parameters.GetOptional("includeStyle", true)
        );
    }

    /// <summary>
    ///     Parameters for the get tab stops operation.
    /// </summary>
    /// <param name="Location">The location type (body, header, footer).</param>
    /// <param name="ParagraphIndex">The paragraph index for body location.</param>
    /// <param name="SectionIndex">The section index.</param>
    /// <param name="AllParagraphs">Whether to get all paragraphs.</param>
    /// <param name="IncludeStyle">Whether to include style tab stops.</param>
    private record GetTabStopsParameters(
        string Location,
        int ParagraphIndex,
        int SectionIndex,
        bool AllParagraphs,
        bool IncludeStyle);
}
