using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for getting run format information in Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
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

            object result;
            if (p.IncludeInherited)
                result = new
                {
                    paragraphIndex = p.ParagraphIndex,
                    runIndex = p.RunIndex.Value,
                    text = run.Text,
                    formatType = "inherited",
                    fontName = font.Name,
                    fontNameAscii = font.NameAscii,
                    fontNameFarEast = font.NameFarEast,
                    fontSize = font.Size,
                    bold = font.Bold,
                    italic = font.Italic,
                    underline = font.Underline.ToString(),
                    strikeThrough = font.StrikeThrough,
                    superscript = font.Superscript,
                    subscript = font.Subscript,
                    color = colorHex,
                    colorName,
                    isAutoColor = font.Color is { IsEmpty: true } or { R: 0, G: 0, B: 0, A: 0 }
                };
            else
                result = new
                {
                    paragraphIndex = p.ParagraphIndex,
                    runIndex = p.RunIndex.Value,
                    text = run.Text,
                    formatType = "explicit",
                    fontName = font.Name,
                    fontNameAscii = font.NameAscii,
                    fontNameFarEast = font.NameFarEast,
                    fontSize = font.Size,
                    bold = font.Bold,
                    italic = font.Italic,
                    underline = font.Underline.ToString(),
                    strikeThrough = font.StrikeThrough,
                    superscript = font.Superscript,
                    subscript = font.Subscript,
                    color = colorHex,
                    colorName,
                    isAutoColor = font.Color is { IsEmpty: true } or { R: 0, G: 0, B: 0, A: 0 }
                };
            return JsonSerializer.Serialize(result, JsonDefaults.Indented);
        }
        else
        {
            var runsList = runs.Select((run, i) =>
            {
                var font = run.Font;
                var colorHex = $"#{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}";
                return new
                {
                    index = i,
                    text = run.Text,
                    fontNameAscii = font.NameAscii,
                    fontNameFarEast = font.NameFarEast,
                    fontSize = font.Size,
                    bold = font.Bold,
                    italic = font.Italic,
                    underline = font.Underline.ToString(),
                    strikeThrough = font.StrikeThrough,
                    superscript = font.Superscript,
                    subscript = font.Subscript,
                    color = colorHex,
                    colorName = WordFormatHelper.GetColorName(font.Color)
                };
            }).ToList();

            var result = new
            {
                paragraphIndex = p.ParagraphIndex,
                count = runs.Count,
                runs = runsList
            };
            return JsonSerializer.Serialize(result, JsonDefaults.Indented);
        }
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
