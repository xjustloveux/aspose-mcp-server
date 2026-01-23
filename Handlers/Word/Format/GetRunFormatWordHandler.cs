using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Format;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for getting run format information in Word documents.
/// </summary>
[ResultType(typeof(GetRunFormatWordResult))]
public class GetRunFormatWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_run_format";

    /// <summary>
    ///     Gets run format information.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    ///     Optional: runIndex, includeInherited
    /// </param>
    /// <returns>A JSON string containing the run format information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetRunFormatParameters(parameters);

        var doc = context.Document;

        var para = WordFormatHelper.GetTargetParagraph(doc, p.ParagraphIndex);
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();

        if (p.RunIndex.HasValue)
        {
            if (p.RunIndex.Value < 0 || p.RunIndex.Value >= runs.Count)
                throw new ArgumentException(
                    $"runIndex {p.RunIndex.Value} is out of range (paragraph #{p.ParagraphIndex} has {runs.Count} Runs, valid range: 0-{runs.Count - 1})");

            var run = runs[p.RunIndex.Value];
            var font = run.Font;
            var colorHex = $"#{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}";
            var colorName = WordFormatHelper.GetColorName(font.Color);

            return new GetRunFormatWordResult
            {
                ParagraphIndex = p.ParagraphIndex,
                RunIndex = p.RunIndex.Value,
                Text = run.Text,
                FormatType = p.IncludeInherited ? "inherited" : "explicit",
                FontName = font.Name,
                FontNameAscii = font.NameAscii,
                FontNameFarEast = font.NameFarEast,
                FontSize = font.Size,
                Bold = font.Bold,
                Italic = font.Italic,
                Underline = font.Underline.ToString(),
                StrikeThrough = font.StrikeThrough,
                Superscript = font.Superscript,
                Subscript = font.Subscript,
                Color = colorHex,
                ColorName = colorName,
                IsAutoColor = font.Color is { IsEmpty: true } or { R: 0, G: 0, B: 0, A: 0 }
            };
        }

        var runsList = runs.Select((run, i) =>
        {
            var font = run.Font;
            var colorHex = $"#{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}";
            return new RunFormatInfo
            {
                Index = i,
                Text = run.Text,
                FontNameAscii = font.NameAscii,
                FontNameFarEast = font.NameFarEast,
                FontSize = font.Size,
                Bold = font.Bold,
                Italic = font.Italic,
                Underline = font.Underline.ToString(),
                StrikeThrough = font.StrikeThrough,
                Superscript = font.Superscript,
                Subscript = font.Subscript,
                Color = colorHex,
                ColorName = WordFormatHelper.GetColorName(font.Color)
            };
        }).ToList();

        return new GetRunFormatAllResult
        {
            ParagraphIndex = p.ParagraphIndex,
            Count = runs.Count,
            Runs = runsList
        };
    }

    private static GetRunFormatParameters ExtractGetRunFormatParameters(OperationParameters parameters)
    {
        return new GetRunFormatParameters(
            parameters.GetOptional("paragraphIndex", 0),
            parameters.GetOptional<int?>("runIndex"),
            parameters.GetOptional("includeInherited", false));
    }

    private sealed record GetRunFormatParameters(
        int ParagraphIndex,
        int? RunIndex,
        bool IncludeInherited);
}
