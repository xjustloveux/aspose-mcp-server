using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.ContentControl;

/// <summary>
///     Handler for editing content control properties in a Word document.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EditWordContentControlHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits the properties of a content control identified by index or tag.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: index (0-based) or tag (to identify the content control)
    ///     Optional: newTag, newTitle, lockContents, lockDeletion
    /// </param>
    /// <returns>Success message with updated properties.</returns>
    /// <exception cref="ArgumentException">Thrown when the content control cannot be found.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditParameters(parameters);

        var doc = context.Document;
        var sdt = FindContentControl(doc, p.Index, p.Tag);

        var changes = new List<string>();

        if (p.NewTag != null)
        {
            sdt.Tag = p.NewTag;
            changes.Add($"tag='{p.NewTag}'");
        }

        if (p.NewTitle != null)
        {
            sdt.Title = p.NewTitle;
            changes.Add($"title='{p.NewTitle}'");
        }

        if (p.LockContents.HasValue)
        {
            sdt.LockContents = p.LockContents.Value;
            changes.Add($"lockContents={p.LockContents.Value}");
        }

        if (p.LockDeletion.HasValue)
        {
            sdt.LockContentControl = p.LockDeletion.Value;
            changes.Add($"lockDeletion={p.LockDeletion.Value}");
        }

        if (changes.Count == 0)
            return new SuccessResult { Message = "No changes specified for the content control." };

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Content control updated: {string.Join(", ", changes)}."
        };
    }

    /// <summary>
    ///     Finds a content control by index or tag.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="index">The 0-based index of the content control.</param>
    /// <param name="tag">The tag identifier of the content control.</param>
    /// <returns>The found StructuredDocumentTag.</returns>
    /// <exception cref="ArgumentException">Thrown when the content control cannot be found.</exception>
    internal static StructuredDocumentTag FindContentControl(Document doc, int? index, string? tag)
    {
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        if (index.HasValue)
        {
            if (index.Value < 0 || index.Value >= sdtNodes.Count)
                throw new ArgumentException(
                    $"Content control index {index.Value} is out of range (document has {sdtNodes.Count} content controls)");

            return (StructuredDocumentTag)sdtNodes[index.Value];
        }

        if (!string.IsNullOrEmpty(tag))
        {
            foreach (var sdt in sdtNodes.Cast<StructuredDocumentTag>())
                if (string.Equals(sdt.Tag, tag, StringComparison.OrdinalIgnoreCase))
                    return sdt;

            throw new ArgumentException($"Content control with tag '{tag}' not found");
        }

        throw new ArgumentException("Either 'index' or 'tag' is required to identify the content control");
    }

    /// <summary>
    ///     Extracts parameters for the edit operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetOptional<int?>("index"),
            parameters.GetOptional<string?>("tag"),
            parameters.GetOptional<string?>("newTag"),
            parameters.GetOptional<string?>("newTitle"),
            parameters.GetOptional<bool?>("lockContents"),
            parameters.GetOptional<bool?>("lockDeletion")
        );
    }

    /// <summary>
    ///     Parameters for the edit content control operation.
    /// </summary>
    /// <param name="Index">The 0-based index of the content control.</param>
    /// <param name="Tag">The tag identifier to find the content control.</param>
    /// <param name="NewTag">The new tag value.</param>
    /// <param name="NewTitle">The new title value.</param>
    /// <param name="LockContents">Whether to lock contents.</param>
    /// <param name="LockDeletion">Whether to lock deletion.</param>
    private sealed record EditParameters(
        int? Index,
        string? Tag,
        string? NewTag,
        string? NewTitle,
        bool? LockContents,
        bool? LockDeletion);
}
