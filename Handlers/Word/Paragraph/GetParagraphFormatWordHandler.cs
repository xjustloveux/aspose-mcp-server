using System.Drawing;
using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for getting paragraph format from Word documents.
/// </summary>
public class GetParagraphFormatWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_format";

    /// <summary>
    ///     Gets detailed formatting information for a specific paragraph.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    ///     Optional: includeRunDetails
    /// </param>
    /// <returns>JSON string containing paragraph formatting information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var includeRunDetails = parameters.GetOptional("includeRunDetails", true);

        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for get_format operation");

        var para = GetParagraph(context.Document, paragraphIndex.Value);
        var resultDict = BuildBasicInfo(para, paragraphIndex.Value);

        AddListFormat(resultDict, para);
        AddBordersInfo(resultDict, para.ParagraphFormat);
        AddBackgroundColor(resultDict, para.ParagraphFormat);
        AddTabStops(resultDict, para.ParagraphFormat);
        AddFontFormat(resultDict, para);
        AddRunDetails(resultDict, para, includeRunDetails);

        return JsonSerializer.Serialize(resultDict, JsonDefaults.Indented);
    }

    private static Aspose.Words.Paragraph GetParagraph(Document doc, int paragraphIndex)
    {
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (paragraphs[paragraphIndex] is not Aspose.Words.Paragraph para)
            throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex}");

        return para;
    }

    private static Dictionary<string, object?> BuildBasicInfo(Aspose.Words.Paragraph para, int paragraphIndex)
    {
        var format = para.ParagraphFormat;
        var text = para.GetText().Trim();

        return new Dictionary<string, object?>
        {
            ["paragraphIndex"] = paragraphIndex,
            ["text"] = text,
            ["textLength"] = text.Length,
            ["runCount"] = para.Runs.Count,
            ["paragraphFormat"] = new
            {
                styleName = format.StyleName,
                alignment = format.Alignment.ToString(),
                leftIndent = Math.Round(format.LeftIndent, 2),
                rightIndent = Math.Round(format.RightIndent, 2),
                firstLineIndent = Math.Round(format.FirstLineIndent, 2),
                spaceBefore = Math.Round(format.SpaceBefore, 2),
                spaceAfter = Math.Round(format.SpaceAfter, 2),
                lineSpacing = Math.Round(format.LineSpacing, 2),
                lineSpacingRule = format.LineSpacingRule.ToString()
            }
        };
    }

    private static void AddListFormat(Dictionary<string, object?> resultDict, Aspose.Words.Paragraph para)
    {
        if (para.ListFormat is not { IsListItem: true }) return;

        resultDict["listFormat"] = new
        {
            isListItem = true,
            listLevel = para.ListFormat.ListLevelNumber,
            listId = para.ListFormat.List?.ListId
        };
    }

    private static void AddBordersInfo(Dictionary<string, object?> resultDict, ParagraphFormat format)
    {
        var borders = new Dictionary<string, object>();

        AddBorderIfPresent(borders, "top", format.Borders.Top);
        AddBorderIfPresent(borders, "bottom", format.Borders.Bottom);
        AddBorderIfPresent(borders, "left", format.Borders.Left);
        AddBorderIfPresent(borders, "right", format.Borders.Right);

        if (borders.Count > 0)
            resultDict["borders"] = borders;
    }

    private static void AddBorderIfPresent(Dictionary<string, object> borders, string name, Border border)
    {
        if (border.LineStyle == LineStyle.None) return;

        borders[name] = new
        {
            lineStyle = border.LineStyle.ToString(),
            lineWidth = border.LineWidth,
            color = border.Color.Name
        };
    }

    private static void AddBackgroundColor(Dictionary<string, object?> resultDict, ParagraphFormat format)
    {
        if (format.Shading.BackgroundPatternColor.ToArgb() == Color.Empty.ToArgb()) return;

        var bgColor = format.Shading.BackgroundPatternColor;
        resultDict["backgroundColor"] = $"#{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";
    }

    private static void AddTabStops(Dictionary<string, object?> resultDict, ParagraphFormat format)
    {
        if (format.TabStops.Count == 0) return;

        List<object> tabStopsList = [];
        for (var i = 0; i < format.TabStops.Count; i++)
        {
            var tab = format.TabStops[i];
            tabStopsList.Add(new
            {
                position = Math.Round(tab.Position, 2),
                alignment = tab.Alignment.ToString(),
                leader = tab.Leader.ToString()
            });
        }

        resultDict["tabStops"] = tabStopsList;
    }

    private static void AddFontFormat(Dictionary<string, object?> resultDict, Aspose.Words.Paragraph para)
    {
        if (para.Runs.Count == 0) return;

        var firstRun = para.Runs[0];
        var fontInfo = BuildFontInfo(firstRun);
        resultDict["fontFormat"] = fontInfo;
    }

    private static Dictionary<string, object?> BuildFontInfo(Run run)
    {
        var fontInfo = new Dictionary<string, object?> { ["fontSize"] = run.Font.Size };

        AddFontNames(fontInfo, run);
        AddFontStyles(fontInfo, run);
        AddFontColors(fontInfo, run);

        return fontInfo;
    }

    private static void AddFontNames(Dictionary<string, object?> fontInfo, Run run)
    {
        if (run.Font.NameAscii != run.Font.NameFarEast)
        {
            fontInfo["fontAscii"] = run.Font.NameAscii;
            fontInfo["fontFarEast"] = run.Font.NameFarEast;
        }
        else
        {
            fontInfo["font"] = run.Font.Name;
        }
    }

    private static void AddFontStyles(Dictionary<string, object?> fontInfo, Run run)
    {
        if (run.Font.Bold) fontInfo["bold"] = true;
        if (run.Font.Italic) fontInfo["italic"] = true;
        if (run.Font.Underline != Underline.None) fontInfo["underline"] = run.Font.Underline.ToString();
        if (run.Font.StrikeThrough) fontInfo["strikethrough"] = true;
        if (run.Font.Superscript) fontInfo["superscript"] = true;
        if (run.Font.Subscript) fontInfo["subscript"] = true;
    }

    private static void AddFontColors(Dictionary<string, object?> fontInfo, Run run)
    {
        if (run.Font.Color.ToArgb() != Color.Empty.ToArgb())
            fontInfo["color"] = $"#{run.Font.Color.R:X2}{run.Font.Color.G:X2}{run.Font.Color.B:X2}";
        if (run.Font.HighlightColor != Color.Empty)
            fontInfo["highlightColor"] = run.Font.HighlightColor.Name;
    }

    private static void AddRunDetails(Dictionary<string, object?> resultDict, Aspose.Words.Paragraph para,
        bool includeRunDetails)
    {
        if (!includeRunDetails || para.Runs.Count <= 1) return;

        List<object> runs = [];
        var displayCount = Math.Min(para.Runs.Count, 10);

        for (var i = 0; i < displayCount; i++) runs.Add(BuildRunInfo(para.Runs[i], i));

        resultDict["runs"] = new { total = para.Runs.Count, displayed = displayCount, details = runs };
    }

    private static Dictionary<string, object?> BuildRunInfo(Run run, int index)
    {
        var runInfo = new Dictionary<string, object?>
        {
            ["index"] = index,
            ["text"] = run.Text.Replace("\r", "\\r").Replace("\n", "\\n"),
            ["fontSize"] = run.Font.Size
        };

        AddFontNames(runInfo, run);

        if (run.Font.Bold) runInfo["bold"] = true;
        if (run.Font.Italic) runInfo["italic"] = true;
        if (run.Font.Underline != Underline.None) runInfo["underline"] = run.Font.Underline.ToString();

        return runInfo;
    }
}
