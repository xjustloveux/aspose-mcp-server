using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Content;

namespace AsposeMcpServer.Handlers.Word.Content;

/// <summary>
///     Handler for getting Word document content as plain text with optional pagination.
/// </summary>
[ResultType(typeof(GetWordContentResult))]
public class GetWordContentHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_content";

    /// <summary>
    ///     Gets document content as plain text with optional pagination.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: maxChars, offset
    /// </param>
    /// <returns>Document content as plain text.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetContentParameters(parameters);

        var document = context.Document;
        var fullText = WordContentHelper.CleanText(document.GetText());
        var totalLength = fullText.Length;

        string content;
        var hasMore = false;
        if (p.Offset >= totalLength)
        {
            content = "";
        }
        else if (p.MaxChars.HasValue)
        {
            var endIndex = Math.Min(p.Offset + p.MaxChars.Value, totalLength);
            content = fullText.Substring(p.Offset, endIndex - p.Offset);
            hasMore = endIndex < totalLength;
        }
        else
        {
            content = p.Offset > 0 ? fullText.Substring(p.Offset) : fullText;
        }

        return new GetWordContentResult
        {
            Content = content,
            TotalLength = totalLength,
            Offset = p.Offset,
            HasMore = hasMore
        };
    }

    /// <summary>
    ///     Extracts parameters for the get content operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetContentParameters ExtractGetContentParameters(OperationParameters parameters)
    {
        var maxChars = parameters.GetOptional<int?>("maxChars");
        var offset = parameters.GetOptional("offset", 0);

        return new GetContentParameters(maxChars, offset);
    }

    /// <summary>
    ///     Parameters for the get content operation.
    /// </summary>
    /// <param name="MaxChars">The maximum number of characters to return.</param>
    /// <param name="Offset">The character offset to start from.</param>
    private sealed record GetContentParameters(int? MaxChars, int Offset);
}
