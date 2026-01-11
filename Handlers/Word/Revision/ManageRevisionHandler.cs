using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Revision;

/// <summary>
///     Handler for managing a specific revision in Word documents (accept or reject).
/// </summary>
public class ManageRevisionHandler : OperationHandlerBase<Document>
{
    private const int MaxRevisionTextLength = 50;

    /// <inheritdoc />
    public override string Operation => "manage";

    /// <summary>
    ///     Manages a specific revision by index (accept or reject).
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: revisionIndex
    ///     Optional: action (default: "accept")
    /// </param>
    /// <returns>Success message with revision details.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when revisionIndex is not provided, is out of range, or action is invalid.
    /// </exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var revisionIndex = parameters.GetRequired<int>("revisionIndex");
        var action = parameters.GetOptional("action", "accept");

        var doc = context.Document;
        var revisionsCount = doc.Revisions.Count;

        if (revisionsCount == 0)
            return Success("Document has no revisions");

        if (revisionIndex < 0 || revisionIndex >= revisionsCount)
            throw new ArgumentException(
                $"revisionIndex must be between 0 and {revisionsCount - 1}, got: {revisionIndex}");

        var revision = doc.Revisions[revisionIndex];
        var revisionType = revision.RevisionType;
        var revisionText = TruncateText(revision.ParentNode?.ToString(SaveFormat.Text)?.Trim() ?? "(none)",
            MaxRevisionTextLength);

        switch (action.ToLowerInvariant())
        {
            case "accept":
                revision.Accept();
                break;
            case "reject":
                revision.Reject();
                break;
            default:
                throw new ArgumentException($"action must be 'accept' or 'reject', got: {action}");
        }

        MarkModified(context);

        var result = $"Revision [{revisionIndex}] {action}ed\n";
        result += $"Type: {revisionType}\n";
        result += $"Text: {revisionText}";
        return Success(result);
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
