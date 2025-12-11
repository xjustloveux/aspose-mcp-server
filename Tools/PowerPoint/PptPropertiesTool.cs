using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint document properties (get, set)
/// Merges: PptGetDocumentPropertiesTool, PptSetDocumentPropertiesTool, PptSetPropertiesTool
/// </summary>
public class PptPropertiesTool : IAsposeTool
{
    public string Description => "Manage PowerPoint document properties: get or set";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'get', 'set'",
                @enum = new[] { "get", "set" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            title = new
            {
                type = "string",
                description = "Title (optional, for set)"
            },
            subject = new
            {
                type = "string",
                description = "Subject (optional, for set)"
            },
            author = new
            {
                type = "string",
                description = "Author (optional, for set)"
            },
            keywords = new
            {
                type = "string",
                description = "Keywords (optional, for set)"
            },
            comments = new
            {
                type = "string",
                description = "Comments (optional, for set)"
            },
            category = new
            {
                type = "string",
                description = "Category (optional, for set)"
            },
            company = new
            {
                type = "string",
                description = "Company (optional, for set)"
            },
            manager = new
            {
                type = "string",
                description = "Manager (optional, for set)"
            },
            customProperties = new
            {
                type = "object",
                description = "Custom properties as key-value pairs (optional, for set)"
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
        using var presentation = new Presentation(path);
        var props = presentation.DocumentProperties;
        var sb = new StringBuilder();

        sb.AppendLine("=== Document Properties ===");
        sb.AppendLine($"Title: {props.Title ?? "(none)"}");
        sb.AppendLine($"Subject: {props.Subject ?? "(none)"}");
        sb.AppendLine($"Author: {props.Author ?? "(none)"}");
        sb.AppendLine($"Keywords: {props.Keywords ?? "(none)"}");
        sb.AppendLine($"Comments: {props.Comments ?? "(none)"}");
        sb.AppendLine($"Category: {props.Category ?? "(none)"}");
        sb.AppendLine($"Company: {props.Company ?? "(none)"}");
        sb.AppendLine($"Manager: {props.Manager ?? "(none)"}");
        sb.AppendLine($"Created: {props.CreatedTime}");
        sb.AppendLine($"Modified: {props.LastSavedTime}");
        sb.AppendLine($"Revision: {props.RevisionNumber}");

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> SetPropertiesAsync(JsonObject? arguments, string path)
    {
        var title = arguments?["title"]?.GetValue<string>();
        var subject = arguments?["subject"]?.GetValue<string>();
        var author = arguments?["author"]?.GetValue<string>();
        var keywords = arguments?["keywords"]?.GetValue<string>();
        var comments = arguments?["comments"]?.GetValue<string>();
        var category = arguments?["category"]?.GetValue<string>();
        var company = arguments?["company"]?.GetValue<string>();
        var manager = arguments?["manager"]?.GetValue<string>();
        var customProps = arguments?["customProperties"]?.AsObject();

        using var presentation = new Presentation(path);
        var props = presentation.DocumentProperties;
        var changes = new List<string>();

        if (!string.IsNullOrEmpty(title))
        {
            props.Title = title;
            changes.Add("Title");
        }
        if (!string.IsNullOrEmpty(subject))
        {
            props.Subject = subject;
            changes.Add("Subject");
        }
        if (!string.IsNullOrEmpty(author))
        {
            props.Author = author;
            changes.Add("Author");
        }
        if (!string.IsNullOrEmpty(keywords))
        {
            props.Keywords = keywords;
            changes.Add("Keywords");
        }
        if (!string.IsNullOrEmpty(comments))
        {
            props.Comments = comments;
            changes.Add("Comments");
        }
        if (!string.IsNullOrEmpty(category))
        {
            props.Category = category;
            changes.Add("Category");
        }
        if (!string.IsNullOrEmpty(company))
        {
            props.Company = company;
            changes.Add("Company");
        }
        if (!string.IsNullOrEmpty(manager))
        {
            props.Manager = manager;
            changes.Add("Manager");
        }

        if (customProps != null)
        {
            foreach (var kvp in customProps)
            {
                props[kvp.Key] = kvp.Value?.GetValue<string>() ?? "";
            }
            changes.Add("CustomProperties");
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Document properties updated: {string.Join(", ", changes)} - {path}");
    }
}

