using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Properties;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Word document properties (get, set)
/// Merges: WordGetDocumentPropertiesTool, WordSetDocumentPropertiesTool, WordSetPropertiesTool
/// </summary>
public class WordPropertiesTool : IAsposeTool
{
    public string Description => @"Get or set Word document properties (metadata). Supports 2 operations: get, set.

Usage examples:
- Get properties: word_properties(operation='get', path='doc.docx')
- Set properties: word_properties(operation='set', path='doc.docx', title='Title', author='Author', subject='Subject')";

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
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for set operation)"
            },
            title = new
            {
                type = "string",
                description = "Document title (optional, for set operation)"
            },
            subject = new
            {
                type = "string",
                description = "Document subject (optional, for set operation)"
            },
            author = new
            {
                type = "string",
                description = "Document author (optional, for set operation)"
            },
            keywords = new
            {
                type = "string",
                description = "Keywords (optional, for set operation)"
            },
            comments = new
            {
                type = "string",
                description = "Comments (optional, for set operation)"
            },
            category = new
            {
                type = "string",
                description = "Document category (optional, for set operation)"
            },
            company = new
            {
                type = "string",
                description = "Company name (optional, for set operation)"
            },
            manager = new
            {
                type = "string",
                description = "Manager name (optional, for set operation)"
            },
            customProperties = new
            {
                type = "object",
                description = "Custom properties as key-value pairs (optional, for set operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

        return operation.ToLower() switch
        {
            "get" => await GetPropertiesAsync(arguments, path),
            "set" => await SetPropertiesAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> GetPropertiesAsync(JsonObject? arguments, string path)
    {
        var doc = new Document(path);
        var props = doc.BuiltInDocumentProperties;
        var customProps = doc.CustomDocumentProperties;
        var sb = new StringBuilder();

        sb.AppendLine("=== Document Properties ===");
        sb.AppendLine();
        sb.AppendLine("Built-in Properties:");
        sb.AppendLine($"  Title: {props.Title ?? "(none)"}");
        sb.AppendLine($"  Subject: {props.Subject ?? "(none)"}");
        sb.AppendLine($"  Author: {props.Author ?? "(none)"}");
        sb.AppendLine($"  Keywords: {props.Keywords ?? "(none)"}");
        sb.AppendLine($"  Comments: {props.Comments ?? "(none)"}");
        sb.AppendLine($"  Category: {props.Category ?? "(none)"}");
        sb.AppendLine($"  Company: {props.Company ?? "(none)"}");
        sb.AppendLine($"  Manager: {props.Manager ?? "(none)"}");
        sb.AppendLine($"  Created: {props.CreatedTime}");
        sb.AppendLine($"  Modified: {props.LastSavedTime}");
        sb.AppendLine($"  Last Saved By: {props.LastSavedBy ?? "(none)"}");
        sb.AppendLine($"  Revision: {props.RevisionNumber}");
        sb.AppendLine($"  Word Count: {props.Words}");
        sb.AppendLine($"  Character Count: {props.Characters}");
        sb.AppendLine($"  Page Count: {props.Pages}");

        if (customProps.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Custom Properties:");
            foreach (DocumentProperty prop in customProps)
            {
                sb.AppendLine($"  {prop.Name}: {prop.Value}");
            }
        }

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> SetPropertiesAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var title = arguments?["title"]?.GetValue<string>();
        var subject = arguments?["subject"]?.GetValue<string>();
        var author = arguments?["author"]?.GetValue<string>();
        var keywords = arguments?["keywords"]?.GetValue<string>();
        var comments = arguments?["comments"]?.GetValue<string>();
        var category = arguments?["category"]?.GetValue<string>();
        var company = arguments?["company"]?.GetValue<string>();
        var manager = arguments?["manager"]?.GetValue<string>();
        var customProps = arguments?["customProperties"]?.AsObject();

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        var props = doc.BuiltInDocumentProperties;

        if (!string.IsNullOrEmpty(title)) props.Title = title;
        if (!string.IsNullOrEmpty(subject)) props.Subject = subject;
        if (!string.IsNullOrEmpty(author)) props.Author = author;
        if (!string.IsNullOrEmpty(keywords)) props.Keywords = keywords;
        if (!string.IsNullOrEmpty(comments)) props.Comments = comments;
        if (!string.IsNullOrEmpty(category)) props.Category = category;
        if (!string.IsNullOrEmpty(company)) props.Company = company;
        if (!string.IsNullOrEmpty(manager)) props.Manager = manager;

        if (customProps != null)
        {
            foreach (var kvp in customProps)
            {
                doc.CustomDocumentProperties.Add(kvp.Key, kvp.Value?.GetValue<string>() ?? "");
            }
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Document properties updated: {outputPath}");
    }
}

