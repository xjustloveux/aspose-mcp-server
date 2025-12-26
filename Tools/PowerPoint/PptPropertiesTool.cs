using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint document properties (get, set)
///     Merges: PptGetDocumentPropertiesTool, PptSetDocumentPropertiesTool, PptSetPropertiesTool
/// </summary>
public class PptPropertiesTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint document properties. Supports 2 operations: get, set.

Usage examples:
- Get properties: ppt_properties(operation='get', path='presentation.pptx')
- Set properties: ppt_properties(operation='set', path='presentation.pptx', title='Title', author='Author', subject='Subject')";

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
                description = "Presentation file path (required for all operations)"
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
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for set operation, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "get" => await GetPropertiesAsync(path),
            "set" => await SetPropertiesAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets presentation properties
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>JSON string with properties</returns>
    private Task<string> GetPropertiesAsync(string path)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);
            var props = presentation.DocumentProperties;

            var result = new
            {
                title = props.Title,
                subject = props.Subject,
                author = props.Author,
                keywords = props.Keywords,
                comments = props.Comments,
                category = props.Category,
                company = props.Company,
                manager = props.Manager,
                createdTime = props.CreatedTime,
                lastSavedTime = props.LastSavedTime,
                revisionNumber = props.RevisionNumber
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Sets presentation properties
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing various property values</param>
    /// <returns>Success message</returns>
    private Task<string> SetPropertiesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var title = ArgumentHelper.GetStringNullable(arguments, "title");
            var subject = ArgumentHelper.GetStringNullable(arguments, "subject");
            var author = ArgumentHelper.GetStringNullable(arguments, "author");
            var keywords = ArgumentHelper.GetStringNullable(arguments, "keywords");
            var comments = ArgumentHelper.GetStringNullable(arguments, "comments");
            var category = ArgumentHelper.GetStringNullable(arguments, "category");
            var company = ArgumentHelper.GetStringNullable(arguments, "company");
            var manager = ArgumentHelper.GetStringNullable(arguments, "manager");
            var customProps = ArgumentHelper.GetObject(arguments, "customProperties", false);

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
                foreach (var kvp in customProps) props[kvp.Key] = kvp.Value?.GetValue<string>() ?? "";
                changes.Add("CustomProperties");
            }

            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Document properties updated: {string.Join(", ", changes)}. Output: {outputPath}";
        });
    }
}