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

        List<WordParagraph> targetParagraphs;
        string locationDesc;

        switch (location.ToLower())
        {
            case "header":
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header != null)
                {
                    var headerParas = header.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
                    targetParagraphs = allParagraphs ? headerParas :
                        headerParas.Count > 0 ? [headerParas[0]] : [];
                    locationDesc = "Header";
                }
                else
                {
                    throw new InvalidOperationException("Header not found");
                }

                break;

            case "footer":
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                {
                    var footerParas = footer.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
                    targetParagraphs = allParagraphs ? footerParas :
                        footerParas.Count > 0 ? [footerParas[0]] : [];
                    locationDesc = "Footer";
                }
                else
                {
                    throw new InvalidOperationException("Footer not found");
                }

                break;

            default:
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
                if (allParagraphs)
                {
                    targetParagraphs = paragraphs;
                }
                else
                {
                    if (paragraphIndex >= paragraphs.Count)
                        throw new ArgumentException(
                            $"Paragraph index {paragraphIndex} out of range (total paragraphs: {paragraphs.Count})");
                    targetParagraphs = [paragraphs[paragraphIndex]];
                }

                locationDesc = allParagraphs ? "Body" : $"Body Paragraph {paragraphIndex}";
                break;
        }

        if (targetParagraphs.Count == 0)
            throw new InvalidOperationException("No target paragraphs found");

        var allTabStops =
            new Dictionary<string, (double position, TabAlignment alignment, TabLeader leader, string source)>();

        for (var paraIdx = 0; paraIdx < targetParagraphs.Count; paraIdx++)
        {
            var para = targetParagraphs[paraIdx];
            var paraSource = allParagraphs ? $"Paragraph {paraIdx}" : "Paragraph";

            var paraTabStops = para.ParagraphFormat.TabStops;
            for (var i = 0; i < paraTabStops.Count; i++)
            {
                var tab = paraTabStops[i];
                var position = Math.Round(tab.Position, 2);
                var key = $"{position}_{tab.Alignment}";
                if (!allTabStops.ContainsKey(key))
                    allTabStops[key] = (position, tab.Alignment, tab.Leader, $"{paraSource} (Custom)");
            }

            if (includeStyle && para.ParagraphFormat.Style != null)
            {
                var currentStyle = para.ParagraphFormat.Style;
                List<Style> styleChain = [];

                while (currentStyle != null)
                {
                    styleChain.Add(currentStyle);
                    if (!string.IsNullOrEmpty(currentStyle.BaseStyleName))
                        try
                        {
                            var baseStyle = para.Document.Styles[currentStyle.BaseStyleName];
                            if (baseStyle != null && !styleChain.Contains(baseStyle))
                                currentStyle = baseStyle;
                            else
                                currentStyle = null;
                        }
                        catch (Exception ex)
                        {
                            currentStyle = null;
                            Console.Error.WriteLine($"[WARN] Error accessing paragraph style: {ex.Message}");
                        }
                    else
                        currentStyle = null;
                }

                foreach (var chainStyle in styleChain)
                    if (chainStyle.ParagraphFormat != null)
                    {
                        var styleTabStops = chainStyle.ParagraphFormat.TabStops;
                        for (var i = 0; i < styleTabStops.Count; i++)
                        {
                            var tab = styleTabStops[i];
                            var position = Math.Round(tab.Position, 2);
                            var key = $"{position}_{tab.Alignment}";

                            if (!allTabStops.ContainsKey(key))
                            {
                                var styleName = chainStyle == para.ParagraphFormat.Style
                                    ? chainStyle.Name
                                    : $"{para.ParagraphFormat.Style.Name} (Base: {chainStyle.Name})";
                                allTabStops[key] = (position, tab.Alignment, tab.Leader,
                                    $"{paraSource} (Style: {styleName})");
                            }
                        }
                    }
            }
        }

        // Build JSON response
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
}
