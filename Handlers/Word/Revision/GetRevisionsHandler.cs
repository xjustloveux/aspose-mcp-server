using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Revision;

namespace AsposeMcpServer.Handlers.Word.Revision;

/// <summary>
///     Handler for getting all revisions from Word documents.
/// </summary>
[ResultType(typeof(GetRevisionsWordResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var revisions = doc.Revisions.ToList();
        List<RevisionInfo> revisionList = [];

        for (var i = 0; i < revisions.Count; i++)
        {
            var revision = revisions[i];
            var text = revision.ParentNode?.ToString(SaveFormat.Text)?.Trim() ?? "(none)";
            var truncatedText = TruncateText(text, MaxRevisionTextLength);

            revisionList.Add(new RevisionInfo
            {
                Index = i,
                Type = revision.RevisionType.ToString(),
                Author = revision.Author,
                Date = revision.DateTime.ToString("yyyy-MM-dd HH:mm:ss"),
                Text = truncatedText
            });
        }

        return new GetRevisionsWordResult
        {
            Count = revisions.Count,
            Revisions = revisionList
        };
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
