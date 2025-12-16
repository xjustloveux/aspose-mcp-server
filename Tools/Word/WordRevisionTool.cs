using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Comparing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordRevisionTool : IAsposeTool
{
    public string Description => @"Manage revisions in Word documents. Supports 5 operations: get_revisions, accept_all, reject_all, manage, compare.

Usage examples:
- Get revisions: word_revision(operation='get_revisions', path='doc.docx')
- Accept all: word_revision(operation='accept_all', path='doc.docx')
- Reject all: word_revision(operation='reject_all', path='doc.docx')
- Manage revision: word_revision(operation='manage', path='doc.docx', revisionIndex=0, action='accept')
- Compare documents: word_revision(operation='compare', path='output.docx', originalPath='original.docx', revisedPath='revised.docx')";

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
- 'compare': Compare two documents (required params: path, originalPath, revisedPath)",
                @enum = new[] { "get_revisions", "accept_all", "reject_all", "manage", "compare" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
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
                description = "Original document file path (for compare)"
            },
            revisedPath = new
            {
                type = "string",
                description = "Revised document file path (for compare)"
            },
            authorName = new
            {
                type = "string",
                description = "Author name for revisions (for compare, default: 'Comparison')"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "get_revisions" => await GetRevisions(arguments),
            "accept_all" => await AcceptAllRevisions(arguments),
            "reject_all" => await RejectAllRevisions(arguments),
            "manage" => await ManageRevisions(arguments),
            "compare" => await CompareDocuments(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Gets all revisions from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path</param>
    /// <returns>Formatted string with all revisions</returns>
    private async Task<string> GetRevisions(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        var doc = new Document(path);
        var sb = new StringBuilder();

        sb.AppendLine("=== Revisions ===");
        sb.AppendLine();

        var revisions = doc.Revisions.ToList();
        for (int i = 0; i < revisions.Count; i++)
        {
            var revision = revisions[i];
            sb.AppendLine($"[{i + 1}] Type: {revision.RevisionType}");
            sb.AppendLine($"    Author: {revision.Author}");
            sb.AppendLine($"    Date: {revision.DateTime}");
            sb.AppendLine($"    Text: {revision.ParentNode?.ToString(SaveFormat.Text)?.Trim() ?? "(none)"}");
            sb.AppendLine();
        }

        sb.AppendLine($"Total Revisions: {revisions.Count}");

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    /// Accepts all revisions in the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> AcceptAllRevisions(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        doc.AcceptAllRevisions();
        doc.Save(outputPath);
        
        return await Task.FromResult($"All revisions accepted: {outputPath}");
    }

    /// <summary>
    /// Rejects all revisions in the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> RejectAllRevisions(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        foreach (var revision in doc.Revisions)
        {
            revision.Reject();
        }
        doc.Save(outputPath);
        
        return await Task.FromResult($"All revisions rejected: {outputPath}");
    }

    /// <summary>
    /// Manages individual revisions (accept/reject specific revisions)
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, revisionIndex, action (accept/reject), optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> ManageRevisions(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var action = ArgumentHelper.GetString(arguments, "action", "accept").ToLowerInvariant();

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        var revisionsCount = doc.Revisions.Count;

        if (revisionsCount == 0)
        {
            if (!string.Equals(path, outputPath, StringComparison.OrdinalIgnoreCase))
            {
                doc.Save(outputPath);
                return await Task.FromResult($"Document has no revisions, saved to: {outputPath}");
            }
            return await Task.FromResult("Document has no revisions");
        }

        switch (action)
        {
            case "accept":
                doc.AcceptAllRevisions();
                break;
            case "reject":
                doc.Revisions.RejectAll();
                break;
            default:
                throw new ArgumentException("action must be 'accept' or 'reject'");
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Processed revisions\nOriginal revisions: {revisionsCount}\nAction: {action}\nOutput: {outputPath}");
    }

    /// <summary>
    /// Compares two documents and shows differences
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, comparePath, optional outputPath</param>
    /// <returns>Success message with comparison result</returns>
    private async Task<string> CompareDocuments(JsonObject? arguments)
    {
        var originalPath = ArgumentHelper.GetString(arguments, "originalPath");
        var revisedPath = ArgumentHelper.GetString(arguments, "revisedPath");
        var outputPath = ArgumentHelper.GetString(arguments, "outputPath");
        var authorName = ArgumentHelper.GetString(arguments, "authorName", "Comparison");

        SecurityHelper.ValidateFilePath(originalPath, "originalPath");
        SecurityHelper.ValidateFilePath(revisedPath, "revisedPath");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var originalDoc = new Document(originalPath);
        var revisedDoc = new Document(revisedPath);

        originalDoc.Compare(revisedDoc, authorName, DateTime.Now);
        originalDoc.Save(outputPath);
        
        return await Task.FromResult($"Comparison document created: {outputPath}");
    }
}

