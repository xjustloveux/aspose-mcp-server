using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfPropertiesTool : IAsposeTool
{
    public string Description => @"Manage document properties in PDF files. Supports 2 operations: get, set.

Usage examples:
- Get properties: pdf_properties(operation='get', path='doc.pdf')
- Set properties: pdf_properties(operation='set', path='doc.pdf', title='Title', author='Author', subject='Subject')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'get': Get document properties (required params: path)
- 'set': Set document properties (required params: path)",
                @enum = new[] { "get", "set" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input for set)"
            },
            title = new
            {
                type = "string",
                description = "Title (for set)"
            },
            author = new
            {
                type = "string",
                description = "Author (for set)"
            },
            subject = new
            {
                type = "string",
                description = "Subject (for set)"
            },
            keywords = new
            {
                type = "string",
                description = "Keywords (for set)"
            },
            creator = new
            {
                type = "string",
                description = "Creator (for set)"
            },
            producer = new
            {
                type = "string",
                description = "Producer (for set)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "get" => await GetProperties(arguments),
            "set" => await SetProperties(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> GetProperties(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var document = new Document(path);
        var metadata = document.Metadata;
        var sb = new StringBuilder();

        sb.AppendLine("=== Document Properties ===");
        sb.AppendLine($"Title: {metadata["Title"] ?? "(none)"}");
        sb.AppendLine($"Author: {metadata["Author"] ?? "(none)"}");
        sb.AppendLine($"Subject: {metadata["Subject"] ?? "(none)"}");
        sb.AppendLine($"Keywords: {metadata["Keywords"] ?? "(none)"}");
        sb.AppendLine($"Creator: {metadata["Creator"] ?? "(none)"}");
        sb.AppendLine($"Producer: {metadata["Producer"] ?? "(none)"}");
        sb.AppendLine($"Creation Date: {metadata["CreationDate"] ?? "(none)"}");
        sb.AppendLine($"Modification Date: {metadata["ModDate"] ?? "(none)"}");
        sb.AppendLine();
        sb.AppendLine($"Total Pages: {document.Pages.Count}");
        sb.AppendLine($"Is Encrypted: {document.IsEncrypted}");
        sb.AppendLine($"Is Linearized: {document.IsLinearized}");

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> SetProperties(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var title = arguments?["title"]?.GetValue<string>();
        var author = arguments?["author"]?.GetValue<string>();
        var subject = arguments?["subject"]?.GetValue<string>();
        var keywords = arguments?["keywords"]?.GetValue<string>();
        var creator = arguments?["creator"]?.GetValue<string>();
        var producer = arguments?["producer"]?.GetValue<string>();

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        var metadata = document.Metadata;

        if (!string.IsNullOrEmpty(title)) metadata["Title"] = title;
        if (!string.IsNullOrEmpty(author)) metadata["Author"] = author;
        if (!string.IsNullOrEmpty(subject)) metadata["Subject"] = subject;
        if (!string.IsNullOrEmpty(keywords)) metadata["Keywords"] = keywords;
        if (!string.IsNullOrEmpty(creator)) metadata["Creator"] = creator;
        if (!string.IsNullOrEmpty(producer)) metadata["Producer"] = producer;

        document.Save(outputPath);
        return await Task.FromResult($"Document properties updated. Output: {outputPath}");
    }
}

