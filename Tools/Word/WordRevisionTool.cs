using System.ComponentModel;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Comparing;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing revision tracking in Word documents
/// </summary>
[McpServerToolType]
public class WordRevisionTool
{
    /// <summary>
    ///     Maximum length of revision text to display in preview
    /// </summary>
    private const int MaxRevisionTextLength = 100;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordRevisionTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    public WordRevisionTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "word_revision")]
    [Description(
        @"Manage revisions in Word documents. Supports 5 operations: get_revisions, accept_all, reject_all, manage, compare.

Usage examples:
- Get revisions: word_revision(operation='get_revisions', path='doc.docx')
- Accept all: word_revision(operation='accept_all', path='doc.docx')
- Reject all: word_revision(operation='reject_all', path='doc.docx')
- Manage specific revision: word_revision(operation='manage', path='doc.docx', revisionIndex=0, action='accept')
- Compare documents: word_revision(operation='compare', path='output.docx', originalPath='original.docx', revisedPath='revised.docx', ignoreFormatting=true)

Notes:
- The 'manage' operation accepts or rejects a specific revision by index (0-based)
- Use 'get_revisions' first to see all revisions and their indices
- Compare operation can optionally ignore formatting and comments changes")]
    public string Execute(
        [Description("Operation: get_revisions, accept_all, reject_all, manage, compare")]
        string operation,
        [Description("Document file path (required if no sessionId for most operations)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional for most operations, required for compare)")]
        string? outputPath = null,
        [Description("Revision index (0-based, required for manage operation)")]
        int? revisionIndex = null,
        [Description("Action for manage operation: accept, reject (default: accept)")]
        string action = "accept",
        [Description("Original document file path (required for compare)")]
        string? originalPath = null,
        [Description("Revised document file path (required for compare)")]
        string? revisedPath = null,
        [Description("Author name for revisions (for compare, default: 'Comparison')")]
        string authorName = "Comparison",
        [Description("Ignore formatting changes in comparison (for compare, default: false)")]
        bool ignoreFormatting = false,
        [Description("Ignore comments in comparison (for compare, default: false)")]
        bool ignoreComments = false)
    {
        return operation.ToLower() switch
        {
            "get_revisions" => GetRevisions(path, sessionId),
            "accept_all" => AcceptAllRevisions(path, sessionId, outputPath),
            "reject_all" => RejectAllRevisions(path, sessionId, outputPath),
            "manage" => ManageRevision(path, sessionId, outputPath, revisionIndex, action),
            "compare" => CompareDocuments(outputPath, originalPath, revisedPath, authorName, ignoreFormatting,
                ignoreComments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets all revisions from the document with truncated text preview.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <returns>A JSON string containing revision information.</returns>
    private string GetRevisions(string? path, string? sessionId)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var doc = ctx.Document;

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

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Truncates text to specified maximum length with ellipsis.
    /// </summary>
    /// <param name="text">The text to truncate.</param>
    /// <param name="maxLength">The maximum length before truncation.</param>
    /// <returns>The truncated text with ellipsis if it exceeds the maximum length.</returns>
    private static string TruncateText(string text, int maxLength)
    {
        if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
            return text;
        return text[..maxLength] + "...";
    }

    /// <summary>
    ///     Accepts all revisions in the document.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string AcceptAllRevisions(string? path, string? sessionId, string? outputPath)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var doc = ctx.Document;

        var count = doc.Revisions.Count;
        doc.AcceptAllRevisions();

        ctx.Save(outputPath);

        var result = $"Accepted {count} revision(s)\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Rejects all revisions in the document.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string RejectAllRevisions(string? path, string? sessionId, string? outputPath)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var doc = ctx.Document;

        var count = doc.Revisions.Count;
        doc.Revisions.RejectAll();

        ctx.Save(outputPath);

        var result = $"Rejected {count} revision(s)\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Manages a specific revision by index (accept or reject).
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="revisionIndex">The zero-based index of the revision to manage.</param>
    /// <param name="action">The action to perform: accept or reject.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when revisionIndex is not provided, is out of range, or action is invalid.</exception>
    private string ManageRevision(string? path, string? sessionId, string? outputPath, int? revisionIndex,
        string action)
    {
        if (!revisionIndex.HasValue)
            throw new ArgumentException("revisionIndex is required for manage operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var doc = ctx.Document;

        var revisionsCount = doc.Revisions.Count;

        if (revisionsCount == 0)
            return "Document has no revisions";

        if (revisionIndex.Value < 0 || revisionIndex.Value >= revisionsCount)
            throw new ArgumentException(
                $"revisionIndex must be between 0 and {revisionsCount - 1}, got: {revisionIndex.Value}");

        var revision = doc.Revisions[revisionIndex.Value];
        var revisionType = revision.RevisionType;
        var revisionText = TruncateText(revision.ParentNode?.ToString(SaveFormat.Text)?.Trim() ?? "(none)", 50);

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

        ctx.Save(outputPath);

        var result = $"Revision [{revisionIndex.Value}] {action}ed\n";
        result += $"Type: {revisionType}\n";
        result += $"Text: {revisionText}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Compares two documents and creates a comparison document showing differences as revisions.
    /// </summary>
    /// <param name="outputPath">The output file path for the comparison document.</param>
    /// <param name="originalPath">The path to the original document.</param>
    /// <param name="revisedPath">The path to the revised document.</param>
    /// <param name="authorName">The author name for the revisions.</param>
    /// <param name="ignoreFormatting">Whether to ignore formatting changes.</param>
    /// <param name="ignoreComments">Whether to ignore comments in comparison.</param>
    /// <returns>A message indicating the comparison result and number of differences found.</returns>
    /// <exception cref="ArgumentException">Thrown when required paths are not provided.</exception>
    private static string CompareDocuments(
        string? outputPath,
        string? originalPath,
        string? revisedPath,
        string authorName,
        bool ignoreFormatting,
        bool ignoreComments)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for compare operation");
        if (string.IsNullOrEmpty(originalPath))
            throw new ArgumentException("originalPath is required for compare operation");
        if (string.IsNullOrEmpty(revisedPath))
            throw new ArgumentException("revisedPath is required for compare operation");

        SecurityHelper.ValidateFilePath(originalPath, "originalPath", true);
        SecurityHelper.ValidateFilePath(revisedPath, "revisedPath", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var originalDoc = new Document(originalPath);
        var revisedDoc = new Document(revisedPath);

        var compareOptions = new CompareOptions
        {
            IgnoreFormatting = ignoreFormatting,
            IgnoreComments = ignoreComments
        };

        originalDoc.Compare(revisedDoc, authorName, DateTime.Now, compareOptions);
        var revisionCount = originalDoc.Revisions.Count;
        originalDoc.Save(outputPath);

        return $"Comparison completed: {revisionCount} difference(s) found\nOutput: {outputPath}";
    }
}