using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

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
        var paragraphIndex = parameters.GetOptional("paragraphIndex", 0);
        var runIndex = parameters.GetOptional<int?>("runIndex");
        var includeInherited = parameters.GetOptional("includeInherited", false);

        var doc = context.Document;

        var para = WordFormatHelper.GetTargetParagraph(doc, paragraphIndex);
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();

        if (runIndex.HasValue)
        {
            if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
                throw new ArgumentException(
                    $"runIndex {runIndex.Value} is out of range (paragraph #{paragraphIndex} has {runs.Count} Runs, valid range: 0-{runs.Count - 1})");

            var run = runs[runIndex.Value];
            var font = run.Font;
            var colorHex = $"#{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}";
            var colorName = WordFormatHelper.GetColorName(font.Color);

            object result;
            if (includeInherited)
                result = new
                {
                    paragraphIndex,
                    runIndex = runIndex.Value,
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
                    paragraphIndex,
                    runIndex = runIndex.Value,
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
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
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
                paragraphIndex,
                count = runs.Count,
                runs = runsList
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
    }
}
