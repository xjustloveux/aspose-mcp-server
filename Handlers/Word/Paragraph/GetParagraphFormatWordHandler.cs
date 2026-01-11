using System.Drawing;
using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

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

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        var para = paragraphs[paragraphIndex.Value] as Aspose.Words.Paragraph;
        if (para == null)
            throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");

        var format = para.ParagraphFormat;
        var text = para.GetText().Trim();

        var resultDict = new Dictionary<string, object?>
        {
            ["paragraphIndex"] = paragraphIndex.Value,
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

        if (para.ListFormat is { IsListItem: true })
            resultDict["listFormat"] = new
            {
                isListItem = true,
                listLevel = para.ListFormat.ListLevelNumber,
                listId = para.ListFormat.List?.ListId
            };

        var borders = new Dictionary<string, object>();
        if (format.Borders.Top.LineStyle != LineStyle.None)
            borders["top"] = new
            {
                lineStyle = format.Borders.Top.LineStyle.ToString(), lineWidth = format.Borders.Top.LineWidth,
                color = format.Borders.Top.Color.Name
            };
        if (format.Borders.Bottom.LineStyle != LineStyle.None)
            borders["bottom"] = new
            {
                lineStyle = format.Borders.Bottom.LineStyle.ToString(), lineWidth = format.Borders.Bottom.LineWidth,
                color = format.Borders.Bottom.Color.Name
            };
        if (format.Borders.Left.LineStyle != LineStyle.None)
            borders["left"] = new
            {
                lineStyle = format.Borders.Left.LineStyle.ToString(), lineWidth = format.Borders.Left.LineWidth,
                color = format.Borders.Left.Color.Name
            };
        if (format.Borders.Right.LineStyle != LineStyle.None)
            borders["right"] = new
            {
                lineStyle = format.Borders.Right.LineStyle.ToString(), lineWidth = format.Borders.Right.LineWidth,
                color = format.Borders.Right.Color.Name
            };
        if (borders.Count > 0)
            resultDict["borders"] = borders;

        if (format.Shading.BackgroundPatternColor.ToArgb() != Color.Empty.ToArgb())
        {
            var bgColor = format.Shading.BackgroundPatternColor;
            resultDict["backgroundColor"] = $"#{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";
        }

        if (format.TabStops.Count > 0)
        {
            List<object> tabStopsList = [];
            for (var i = 0; i < format.TabStops.Count; i++)
            {
                var tab = format.TabStops[i];
                tabStopsList.Add(new
                {
                    position = Math.Round(tab.Position, 2), alignment = tab.Alignment.ToString(),
                    leader = tab.Leader.ToString()
                });
            }

            resultDict["tabStops"] = tabStopsList;
        }

        if (para.Runs.Count > 0)
        {
            var firstRun = para.Runs[0];
            var fontInfo = new Dictionary<string, object?> { ["fontSize"] = firstRun.Font.Size };

            if (firstRun.Font.NameAscii != firstRun.Font.NameFarEast)
            {
                fontInfo["fontAscii"] = firstRun.Font.NameAscii;
                fontInfo["fontFarEast"] = firstRun.Font.NameFarEast;
            }
            else
            {
                fontInfo["font"] = firstRun.Font.Name;
            }

            if (firstRun.Font.Bold) fontInfo["bold"] = true;
            if (firstRun.Font.Italic) fontInfo["italic"] = true;
            if (firstRun.Font.Underline != Underline.None) fontInfo["underline"] = firstRun.Font.Underline.ToString();
            if (firstRun.Font.StrikeThrough) fontInfo["strikethrough"] = true;
            if (firstRun.Font.Superscript) fontInfo["superscript"] = true;
            if (firstRun.Font.Subscript) fontInfo["subscript"] = true;
            if (firstRun.Font.Color.ToArgb() != Color.Empty.ToArgb())
                fontInfo["color"] = $"#{firstRun.Font.Color.R:X2}{firstRun.Font.Color.G:X2}{firstRun.Font.Color.B:X2}";
            if (firstRun.Font.HighlightColor != Color.Empty)
                fontInfo["highlightColor"] = firstRun.Font.HighlightColor.Name;

            resultDict["fontFormat"] = fontInfo;
        }

        if (includeRunDetails && para.Runs.Count > 1)
        {
            List<object> runs = [];
            for (var i = 0; i < Math.Min(para.Runs.Count, 10); i++)
            {
                var run = para.Runs[i];
                var runInfo = new Dictionary<string, object?>
                {
                    ["index"] = i,
                    ["text"] = run.Text.Replace("\r", "\\r").Replace("\n", "\\n"),
                    ["fontSize"] = run.Font.Size
                };

                if (run.Font.NameAscii != run.Font.NameFarEast)
                {
                    runInfo["fontAscii"] = run.Font.NameAscii;
                    runInfo["fontFarEast"] = run.Font.NameFarEast;
                }
                else
                {
                    runInfo["font"] = run.Font.Name;
                }

                if (run.Font.Bold) runInfo["bold"] = true;
                if (run.Font.Italic) runInfo["italic"] = true;
                if (run.Font.Underline != Underline.None) runInfo["underline"] = run.Font.Underline.ToString();

                runs.Add(runInfo);
            }

            resultDict["runs"] = new
                { total = para.Runs.Count, displayed = Math.Min(para.Runs.Count, 10), details = runs };
        }

        return JsonSerializer.Serialize(resultDict, new JsonSerializerOptions { WriteIndented = true });
    }
}
