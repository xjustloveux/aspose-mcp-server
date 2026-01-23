using System.Drawing;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for getting paragraph format from Word documents.
/// </summary>
[ResultType(typeof(GetParagraphFormatWordResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParagraphFormatParameters(parameters);

        if (!getParams.ParagraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for get_format operation");

        var para = GetParagraph(context.Document, getParams.ParagraphIndex.Value);

        return BuildResult(para, getParams.ParagraphIndex.Value, getParams.IncludeRunDetails);
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

    private static GetParagraphFormatWordResult BuildResult(Aspose.Words.Paragraph para, int paragraphIndex,
        bool includeRunDetails)
    {
        var format = para.ParagraphFormat;
        var text = para.GetText().Trim();

        return new GetParagraphFormatWordResult
        {
            ParagraphIndex = paragraphIndex,
            Text = text,
            TextLength = text.Length,
            RunCount = para.Runs.Count,
            ParagraphFormat = new ParagraphFormatInfo
            {
                StyleName = format.StyleName,
                Alignment = format.Alignment.ToString(),
                LeftIndent = Math.Round(format.LeftIndent, 2),
                RightIndent = Math.Round(format.RightIndent, 2),
                FirstLineIndent = Math.Round(format.FirstLineIndent, 2),
                SpaceBefore = Math.Round(format.SpaceBefore, 2),
                SpaceAfter = Math.Round(format.SpaceAfter, 2),
                LineSpacing = Math.Round(format.LineSpacing, 2),
                LineSpacingRule = format.LineSpacingRule.ToString()
            },
            ListFormat = BuildListFormat(para),
            Borders = BuildBorders(format),
            BackgroundColor = BuildBackgroundColor(format),
            TabStops = BuildTabStops(format),
            FontFormat = BuildFontFormat(para),
            Runs = BuildRunDetails(para, includeRunDetails)
        };
    }

    private static ParagraphListFormatInfo? BuildListFormat(Aspose.Words.Paragraph para)
    {
        if (para.ListFormat is not { IsListItem: true }) return null;

        return new ParagraphListFormatInfo
        {
            IsListItem = true,
            ListLevel = para.ListFormat.ListLevelNumber,
            ListId = para.ListFormat.List?.ListId
        };
    }

    private static Dictionary<string, BorderInfo>? BuildBorders(ParagraphFormat format)
    {
        var borders = new Dictionary<string, BorderInfo>();

        AddBorderIfPresent(borders, "top", format.Borders.Top);
        AddBorderIfPresent(borders, "bottom", format.Borders.Bottom);
        AddBorderIfPresent(borders, "left", format.Borders.Left);
        AddBorderIfPresent(borders, "right", format.Borders.Right);

        return borders.Count > 0 ? borders : null;
    }

    private static void AddBorderIfPresent(Dictionary<string, BorderInfo> borders, string name, Border border)
    {
        if (border.LineStyle == LineStyle.None) return;

        borders[name] = new BorderInfo
        {
            LineStyle = border.LineStyle.ToString(),
            LineWidth = border.LineWidth,
            Color = border.Color.Name
        };
    }

    private static string? BuildBackgroundColor(ParagraphFormat format)
    {
        if (format.Shading.BackgroundPatternColor.ToArgb() == Color.Empty.ToArgb()) return null;

        var bgColor = format.Shading.BackgroundPatternColor;
        return $"#{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";
    }

    private static List<ParagraphTabStopInfo>? BuildTabStops(ParagraphFormat format)
    {
        if (format.TabStops.Count == 0) return null;

        List<ParagraphTabStopInfo> tabStopsList = [];
        for (var i = 0; i < format.TabStops.Count; i++)
        {
            var tab = format.TabStops[i];
            tabStopsList.Add(new ParagraphTabStopInfo
            {
                Position = Math.Round(tab.Position, 2),
                Alignment = tab.Alignment.ToString(),
                Leader = tab.Leader.ToString()
            });
        }

        return tabStopsList;
    }

    private static FontFormatInfo? BuildFontFormat(Aspose.Words.Paragraph para)
    {
        if (para.Runs.Count == 0) return null;

        var firstRun = para.Runs[0];
        var font = firstRun.Font;

        return new FontFormatInfo
        {
            FontSize = font.Size,
            Font = font.NameAscii == font.NameFarEast ? font.Name : null,
            FontAscii = font.NameAscii != font.NameFarEast ? font.NameAscii : null,
            FontFarEast = font.NameAscii != font.NameFarEast ? font.NameFarEast : null,
            Bold = font.Bold,
            Italic = font.Italic,
            Underline = font.Underline != Underline.None ? font.Underline.ToString() : null,
            Strikethrough = font.StrikeThrough,
            Superscript = font.Superscript,
            Subscript = font.Subscript,
            Color = font.Color.ToArgb() != Color.Empty.ToArgb()
                ? $"#{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}"
                : null,
            HighlightColor = font.HighlightColor != Color.Empty ? font.HighlightColor.Name : null
        };
    }

    private static RunDetailsInfo? BuildRunDetails(Aspose.Words.Paragraph para, bool includeRunDetails)
    {
        if (!includeRunDetails || para.Runs.Count <= 1) return null;

        var displayCount = Math.Min(para.Runs.Count, 10);
        List<RunDetailInfo> runs = [];

        for (var i = 0; i < displayCount; i++)
            runs.Add(BuildRunInfo(para.Runs[i], i));

        return new RunDetailsInfo
        {
            Total = para.Runs.Count,
            Displayed = displayCount,
            Details = runs
        };
    }

    private static RunDetailInfo BuildRunInfo(Run run, int index)
    {
        var font = run.Font;
        return new RunDetailInfo
        {
            Index = index,
            Text = run.Text.Replace("\r", "\\r").Replace("\n", "\\n"),
            FontSize = font.Size,
            Font = font.NameAscii == font.NameFarEast ? font.Name : null,
            FontAscii = font.NameAscii != font.NameFarEast ? font.NameAscii : null,
            FontFarEast = font.NameAscii != font.NameFarEast ? font.NameFarEast : null,
            Bold = font.Bold,
            Italic = font.Italic,
            Underline = font.Underline != Underline.None ? font.Underline.ToString() : null
        };
    }

    /// <summary>
    ///     Extracts get paragraph format parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get paragraph format parameters.</returns>
    private static GetParagraphFormatParameters ExtractGetParagraphFormatParameters(OperationParameters parameters)
    {
        return new GetParagraphFormatParameters(
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional("includeRunDetails", true)
        );
    }

    /// <summary>
    ///     Record to hold get paragraph format parameters.
    /// </summary>
    /// <param name="ParagraphIndex">The paragraph index to get format for.</param>
    /// <param name="IncludeRunDetails">Whether to include run details.</param>
    private sealed record GetParagraphFormatParameters(int? ParagraphIndex, bool IncludeRunDetails);
}
