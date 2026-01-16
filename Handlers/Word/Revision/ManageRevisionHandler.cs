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
        var p = ExtractManageRevisionParameters(parameters);

        var doc = context.Document;
        var revisionsCount = doc.Revisions.Count;

        if (revisionsCount == 0)
            return Success("Document has no revisions");

        if (p.RevisionIndex < 0 || p.RevisionIndex >= revisionsCount)
            throw new ArgumentException(
                $"revisionIndex must be between 0 and {revisionsCount - 1}, got: {p.RevisionIndex}");

        var revision = doc.Revisions[p.RevisionIndex];
        var revisionType = revision.RevisionType;
        var revisionText = TruncateText(revision.ParentNode?.ToString(SaveFormat.Text)?.Trim() ?? "(none)",
            MaxRevisionTextLength);

        switch (p.Action.ToLowerInvariant())
        {
            case "accept":
                revision.Accept();
                break;
            case "reject":
                revision.Reject();
                break;
            default:
                throw new ArgumentException($"action must be 'accept' or 'reject', got: {p.Action}");
        }

        MarkModified(context);

        var result = $"Revision [{p.RevisionIndex}] {p.Action}ed\n";
        result += $"Type: {revisionType}\n";
        result += $"Text: {revisionText}";
        return Success(result);
    }

    private static ManageRevisionParameters ExtractManageRevisionParameters(OperationParameters parameters)
    {
        return new ManageRevisionParameters(
            parameters.GetRequired<int>("revisionIndex"),
            parameters.GetOptional("action", "accept"));
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

    private record ManageRevisionParameters(int RevisionIndex, string Action);
}
