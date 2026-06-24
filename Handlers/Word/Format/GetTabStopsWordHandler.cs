using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Format;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Represents the target paragraphs and location description for tab stop operations.
/// </summary>
/// <param name="Paragraphs">The list of target paragraphs.</param>
/// <param name="LocationDescription">The description of the location.</param>
internal record TabStopTarget(List<WordParagraph> Paragraphs, string LocationDescription);

/// <summary>
///     Handler for getting tab stops in Word documents.
/// </summary>
[ResultType(typeof(GetTabStopsWordResult))]
public class GetTabStopsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_tabs";

    /// <summary>
    ///     Gets tab stops for a paragraph.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    ///     Optional: location, sectionIndex, allParagraphs, includeStyle
    /// </param>
    /// <returns>A JSON string containing the tab stops information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetTabStopsParameters(parameters);

        var doc = context.Document;

        var tabStopTarget =
            GetTargetParagraphs(doc, p.Location, p.SectionIndex, p.ParagraphIndex, p.AllParagraphs);
        var targetParagraphs = tabStopTarget.Paragraphs;
        var locationDesc = tabStopTarget.LocationDescription;

        if (targetParagraphs.Count == 0)
            throw new InvalidOperationException("No target paragraphs found");

        var allTabStops = CollectTabStops(targetParagraphs, p.AllParagraphs, p.IncludeStyle);

        var tabStopsList = allTabStops
            .OrderBy(x => x.Value.position)
            .Select(kvp => new TabStopDetailInfo
            {
                PositionPt = kvp.Value.position,
                PositionCm = Math.Round(kvp.Value.position / 28.35, 2),
                Alignment = kvp.Value.alignment.ToString(),
                Leader = kvp.Value.leader.ToString(),
                Source = kvp.Value.source
            })
            .ToList();

        return new GetTabStopsWordResult
        {
            Location = p.Location,
            LocationDescription = locationDesc,
            SectionIndex = p.SectionIndex,
            ParagraphIndex = p is { Location: "body", AllParagraphs: false } ? p.ParagraphIndex : null,
            AllParagraphs = p.AllParagraphs,
            IncludeStyle = p.IncludeStyle,
            ParagraphCount = targetParagraphs.Count,
            Count = allTabStops.Count,
            TabStops = tabStopsList
        };
    }

    /// <summary>
    ///     Gets the target paragraphs based on location and settings.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="location">The location type (body, header, footer).</param>
    /// <param name="sectionIndex">The section index.</param>
    /// <param name="paragraphIndex">The paragraph index for body location.</param>
    /// <param name="allParagraphs">Whether to get all paragraphs.</param>
    /// <returns>A TabStopTarget containing the list of paragraphs and location description.</returns>
    private static TabStopTarget GetTargetParagraphs(Document doc, string location, int sectionIndex,
        int paragraphIndex, bool allParagraphs)
    {
        var (storyType, desc) = location.ToLower() switch
        {
            "header" => (StoryTypes.Header, "Header"),
            "footer" => (StoryTypes.Footer, "Footer"),
            _ => (StoryTypes.Body, "Body")
        };

        if (storyType is StoryTypes.Header or StoryTypes.Footer
            && sectionIndex >= 0 && sectionIndex < doc.Sections.Count)
        {
            var hfType = storyType == StoryTypes.Header
                ? HeaderFooterType.HeaderPrimary
                : HeaderFooterType.FooterPrimary;
            if (doc.Sections[sectionIndex].HeadersFooters[hfType] == null)
                throw new InvalidOperationException($"{desc} not found");
        }

        if (allParagraphs)
            return new TabStopTarget(
                ParagraphResolver.GetStoryParagraphs(doc, new ParagraphAddress(0, storyType, sectionIndex)), desc);

        var index = storyType == StoryTypes.Body ? paragraphIndex : 0;
        var para = ParagraphResolver.Resolve(doc, new ParagraphAddress(index, storyType, sectionIndex)).Paragraph;
        var singleDesc = storyType == StoryTypes.Body ? $"Body Paragraph {paragraphIndex}" : desc;
        return new TabStopTarget([para], singleDesc);
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
    private sealed record GetTabStopsParameters(
        string Location,
        int ParagraphIndex,
        int SectionIndex,
        bool AllParagraphs,
        bool IncludeStyle);
}
