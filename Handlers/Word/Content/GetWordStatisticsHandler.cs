using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Content;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Content;

/// <summary>
///     Handler for getting Word document statistics.
/// </summary>
[ResultType(typeof(GetWordStatisticsResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetStatisticsParameters(parameters);

        var document = context.Document;
        document.UpdateWordCount();

        var stats = document.BuiltInDocumentProperties;

        var tables = document.GetChildNodes(NodeType.Table, true);
        var shapes = document.GetChildNodes(NodeType.Shape, true).Cast<WordShape>().ToList();
        var images = shapes.Count(s => s.HasImage);

        return new GetWordStatisticsResult
        {
            Pages = stats.Pages,
            Words = stats.Words,
            Characters = stats.Characters,
            CharactersWithSpaces = stats.CharactersWithSpaces,
            Paragraphs = stats.Paragraphs,
            Lines = stats.Lines,
            Footnotes = p.IncludeFootnotes ? document.GetChildNodes(NodeType.Footnote, true).Count : null,
            FootnotesIncluded = p.IncludeFootnotes,
            Tables = tables.Count,
            Images = images,
            Shapes = shapes.Count,
            StatisticsUpdated = true
        };
    }

    /// <summary>
    ///     Extracts parameters for the get statistics operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetStatisticsParameters ExtractGetStatisticsParameters(OperationParameters parameters)
    {
        var includeFootnotes = parameters.GetOptional("includeFootnotes", true);

        return new GetStatisticsParameters(includeFootnotes);
    }

    /// <summary>
    ///     Parameters for the get statistics operation.
    /// </summary>
    /// <param name="IncludeFootnotes">Whether to include footnote count in the statistics.</param>
    private sealed record GetStatisticsParameters(bool IncludeFootnotes);
}
