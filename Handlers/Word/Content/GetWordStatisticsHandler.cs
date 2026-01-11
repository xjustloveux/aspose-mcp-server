using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Content;

/// <summary>
///     Handler for getting Word document statistics.
/// </summary>
public class GetWordStatisticsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_statistics";

    /// <summary>
    ///     Gets document statistics including word count, page count, and element counts.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: includeFootnotes (default: true)
    /// </param>
    /// <returns>JSON string containing document statistics.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var includeFootnotes = parameters.GetOptional("includeFootnotes", true);

        var document = context.Document;
        document.UpdateWordCount();

        var stats = document.BuiltInDocumentProperties;

        var tables = document.GetChildNodes(NodeType.Table, true);
        var shapes = document.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().ToList();
        var images = shapes.Count(s => s.HasImage);

        var result = new
        {
            pages = stats.Pages,
            words = stats.Words,
            characters = stats.Characters,
            charactersWithSpaces = stats.CharactersWithSpaces,
            paragraphs = stats.Paragraphs,
            lines = stats.Lines,
            footnotes = includeFootnotes ? document.GetChildNodes(NodeType.Footnote, true).Count : (int?)null,
            footnotesIncluded = includeFootnotes,
            tables = tables.Count,
            images,
            shapes = shapes.Count,
            statisticsUpdated = true
        };

        return JsonResult(result);
    }
}
