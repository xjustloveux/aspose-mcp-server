using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Revision;

/// <summary>
///     Handler for getting all revisions from Word documents.
/// </summary>
public class GetRevisionsHandler : OperationHandlerBase<Document>
{
    private const int MaxRevisionTextLength = 100;

    /// <inheritdoc />
    public override string Operation => "get_revisions";

    /// <summary>
    ///     Gets all revisions from the document with truncated text preview.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing revision information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var revisions = doc.Revisions.ToList();
        List<object> revisionList = [];

        for (var i = 0; i < revisions.Count; i++)
        {
            var revision = revisions[i];
            var text = revision.ParentNode?.ToString(SaveFormat.Text)?.Trim() ?? "(none)";
            var truncatedText = TruncateText(text, MaxRevisionTextLength);

            revisionList.Add(new
            {
                index = i,
                type = revision.RevisionType.ToString(),
                author = revision.Author,
                date = revision.DateTime.ToString("yyyy-MM-dd HH:mm:ss"),
                text = truncatedText
            });
        }

        var result = new
        {
            count = revisions.Count,
            revisions = revisionList
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Truncates text to the specified maximum length with ellipsis.
    /// </summary>
    /// <param name="text">The text to truncate.</param>
    /// <param name="maxLength">The maximum length including ellipsis.</param>
    /// <returns>The truncated text with ellipsis if needed.</returns>
    private static string TruncateText(string text, int maxLength)
    {
        if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
            return text;
        return text[..maxLength] + "...";
    }
}
