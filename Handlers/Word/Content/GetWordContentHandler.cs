using System.Text;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Content;

/// <summary>
///     Handler for getting Word document content as plain text with optional pagination.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var maxChars = parameters.GetOptional<int?>("maxChars");
        var offset = parameters.GetOptional("offset", 0);

        var document = context.Document;
        var fullText = WordContentHelper.CleanText(document.GetText());
        var totalLength = fullText.Length;

        string content;
        var hasMore = false;
        if (offset >= totalLength)
        {
            content = "";
        }
        else if (maxChars.HasValue)
        {
            var endIndex = Math.Min(offset + maxChars.Value, totalLength);
            content = fullText.Substring(offset, endIndex - offset);
            hasMore = endIndex < totalLength;
        }
        else
        {
            content = offset > 0 ? fullText.Substring(offset) : fullText;
        }

        var sb = new StringBuilder();
        sb.AppendLine("=== Document Content ===");
        if (maxChars.HasValue || offset > 0)
        {
            sb.AppendLine($"[Showing chars {offset} to {offset + content.Length} of {totalLength}]");
            if (hasMore)
                sb.AppendLine($"[More content available, use offset={offset + content.Length} to continue]");
        }

        sb.AppendLine(content);
        return sb.ToString();
    }
}
