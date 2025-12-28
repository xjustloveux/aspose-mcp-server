using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Comparing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing revision tracking in Word documents
/// </summary>
public class WordRevisionTool : IAsposeTool
{
    private const int MaxRevisionTextLength = 100;

    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
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
- Compare operation can optionally ignore formatting and comments changes";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'get_revisions': Get all revisions (required params: path)
- 'accept_all': Accept all revisions (required params: path)
- 'reject_all': Reject all revisions (required params: path)
- 'manage': Manage a specific revision (required params: path, revisionIndex, action)
- 'compare': Compare two documents (required params: outputPath, originalPath, revisedPath)",
                @enum = new[] { "get_revisions", "accept_all", "reject_all", "manage", "compare" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for get_revisions, accept_all, reject_all, manage)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional for most operations, required for compare)"
            },
            revisionIndex = new
            {
                type = "number",
                description = "Revision index (0-based, required for manage operation)"
            },
            action = new
            {
                type = "string",
                description = "Action for manage operation: accept, reject (default: accept)",
                @enum = new[] { "accept", "reject" }
            },
            originalPath = new
            {
                type = "string",
                description = "Original document file path (required for compare)"
            },
            revisedPath = new
            {
                type = "string",
                description = "Revised document file path (required for compare)"
            },
            authorName = new
            {
                type = "string",
                description = "Author name for revisions (for compare, default: 'Comparison')"
            },
            ignoreFormatting = new
            {
                type = "boolean",
                description = "Ignore formatting changes in comparison (for compare, default: false)"
            },
            ignoreComments = new
            {
                type = "boolean",
                description = "Ignore comments in comparison (for compare, default: false)"
            }
        },
        required = new[] { "operation" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "get_revisions" => await GetRevisionsAsync(arguments),
            "accept_all" => await AcceptAllRevisionsAsync(arguments),
            "reject_all" => await RejectAllRevisionsAsync(arguments),
            "manage" => await ManageRevisionAsync(arguments),
            "compare" => await CompareDocumentsAsync(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets all revisions from the document with truncated text preview
    /// </summary>
    /// <param name="arguments">JSON arguments containing path</param>
    /// <returns>JSON formatted string with all revisions including index, type, author, date, and truncated text</returns>
    private Task<string> GetRevisionsAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);

            var doc = new Document(path);

            var revisions = doc.Revisions.ToList();
            var revisionList = new List<object>();

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
        });
    }

    /// <summary>
    ///     Truncates text to specified maximum length with ellipsis
    /// </summary>
    /// <param name="text">Text to truncate</param>
    /// <param name="maxLength">Maximum length before truncation</param>
    /// <returns>Truncated text with "..." if exceeds maxLength, otherwise original text</returns>
    private static string TruncateText(string text, int maxLength)
    {
        if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
            return text;
        return text[..maxLength] + "...";
    }

    /// <summary>
    ///     Accepts all revisions in the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional outputPath</param>
    /// <returns>Success message with revision count and output path</returns>
    private Task<string> AcceptAllRevisionsAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);
            var count = doc.Revisions.Count;
            doc.AcceptAllRevisions();
            doc.Save(outputPath);

            return $"Accepted {count} revision(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Rejects all revisions in the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional outputPath</param>
    /// <returns>Success message with revision count and output path</returns>
    private Task<string> RejectAllRevisionsAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);
            var count = doc.Revisions.Count;
            doc.Revisions.RejectAll();
            doc.Save(outputPath);

            return $"Rejected {count} revision(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Manages a specific revision by index (accept or reject)
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, revisionIndex, action (accept/reject), optional outputPath</param>
    /// <returns>Success message with revision details and output path</returns>
    /// <exception cref="ArgumentException">Thrown when revisionIndex is out of range or action is invalid</exception>
    private Task<string> ManageRevisionAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var revisionIndex = ArgumentHelper.GetInt(arguments, "revisionIndex");
            var action = ArgumentHelper.GetString(arguments, "action", "accept").ToLowerInvariant();
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);
            var revisionsCount = doc.Revisions.Count;

            if (revisionsCount == 0)
                return "Document has no revisions";

            if (revisionIndex < 0 || revisionIndex >= revisionsCount)
                throw new ArgumentException(
                    $"revisionIndex must be between 0 and {revisionsCount - 1}, got: {revisionIndex}");

            var revision = doc.Revisions[revisionIndex];
            var revisionType = revision.RevisionType;
            var revisionText = TruncateText(revision.ParentNode?.ToString(SaveFormat.Text)?.Trim() ?? "(none)", 50);

            switch (action)
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

            doc.Save(outputPath);
            return
                $"Revision [{revisionIndex}] {action}ed\nType: {revisionType}\nText: {revisionText}\nOutput: {outputPath}";
        });
    }

    /// <summary>
    ///     Compares two documents and creates a comparison document showing differences as revisions
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing originalPath, revisedPath, outputPath, authorName, ignoreFormatting,
    ///     ignoreComments
    /// </param>
    /// <returns>Success message with revision count and output path</returns>
    private Task<string> CompareDocumentsAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var originalPath = ArgumentHelper.GetString(arguments, "originalPath");
            var revisedPath = ArgumentHelper.GetString(arguments, "revisedPath");
            var outputPath = ArgumentHelper.GetString(arguments, "outputPath");
            var authorName = ArgumentHelper.GetString(arguments, "authorName", "Comparison");
            var ignoreFormatting = ArgumentHelper.GetBool(arguments, "ignoreFormatting", false);
            var ignoreComments = ArgumentHelper.GetBool(arguments, "ignoreComments", false);

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
        });
    }
}