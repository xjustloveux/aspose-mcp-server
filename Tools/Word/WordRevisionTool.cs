using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Comparing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordRevisionTool : IAsposeTool
{
    public string Description => "Manage revisions in Word documents (get, accept all, reject all, manage, compare documents)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation: get_revisions, accept_all, reject_all, manage, compare",
                @enum = new[] { "get_revisions", "accept_all", "reject_all", "manage", "compare" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
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
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

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

    private async Task<string> GetRevisions(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

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

    private async Task<string> AcceptAllRevisions(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        doc.AcceptAllRevisions();
        doc.Save(outputPath);
        
        return await Task.FromResult($"All revisions accepted: {outputPath}");
    }

    private async Task<string> RejectAllRevisions(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        foreach (var revision in doc.Revisions)
        {
            revision.Reject();
        }
        doc.Save(outputPath);
        
        return await Task.FromResult($"All revisions rejected: {outputPath}");
    }

    private async Task<string> ManageRevisions(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var action = arguments?["action"]?.GetValue<string>()?.ToLowerInvariant() ?? "accept";

        SecurityHelper.ValidateFilePath(path, "path");
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

    private async Task<string> CompareDocuments(JsonObject? arguments)
    {
        var originalPath = arguments?["originalPath"]?.GetValue<string>() ?? throw new ArgumentException("originalPath is required");
        var revisedPath = arguments?["revisedPath"]?.GetValue<string>() ?? throw new ArgumentException("revisedPath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var authorName = arguments?["authorName"]?.GetValue<string>() ?? "Comparison";

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

