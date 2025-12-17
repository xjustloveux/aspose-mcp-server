using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing document properties in PDF files (get, set)
/// </summary>
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
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "get" => await GetProperties(arguments),
            "set" => await SetProperties(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets PDF properties
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>Formatted string with properties</returns>
    private async Task<string> GetProperties(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);

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

    /// <summary>
    ///     Sets PDF properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing various property values, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> SetProperties(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var title = ArgumentHelper.GetStringNullable(arguments, "title");
        var author = ArgumentHelper.GetStringNullable(arguments, "author");
        var subject = ArgumentHelper.GetStringNullable(arguments, "subject");
        var keywords = ArgumentHelper.GetStringNullable(arguments, "keywords");
        var creator = ArgumentHelper.GetStringNullable(arguments, "creator");
        var producer = ArgumentHelper.GetStringNullable(arguments, "producer");

        SecurityHelper.ValidateFilePath(path);
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