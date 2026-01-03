using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint document properties (get, set).
///     Uses IPresentationInfo for efficient property reading without loading entire presentation.
/// </summary>
[McpServerToolType]
public class PptPropertiesTool
{
    /// <summary>
    ///     JSON serializer options for consistent output formatting.
    /// </summary>
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptPropertiesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    public PptPropertiesTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_properties")]
    [Description(@"Manage PowerPoint document properties. Supports 2 operations: get, set.

Warning: If outputPath is not provided for 'set' operation, the original file will be overwritten.
Note: Custom properties support multiple types: string, int, double, bool, DateTime (ISO format).

Usage examples:
- Get properties: ppt_properties(operation='get', path='presentation.pptx')
- Set properties: ppt_properties(operation='set', path='presentation.pptx', title='Title', author='Author')
- Set custom properties: ppt_properties(operation='set', path='presentation.pptx', customProperties={'Count': 42, 'IsPublished': true})")]
    public string Execute(
        [Description("Operation: get, set")] string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Title (optional, for set)")]
        string? title = null,
        [Description("Subject (optional, for set)")]
        string? subject = null,
        [Description("Author (optional, for set)")]
        string? author = null,
        [Description("Keywords (optional, for set)")]
        string? keywords = null,
        [Description("Comments (optional, for set)")]
        string? comments = null,
        [Description("Category (optional, for set)")]
        string? category = null,
        [Description("Company (optional, for set)")]
        string? company = null,
        [Description("Manager (optional, for set)")]
        string? manager = null,
        [Description(
            "Custom properties as key-value pairs. Supports: string, int, double, bool, DateTime (ISO format).")]
        Dictionary<string, object>? customProperties = null)
    {
        return operation.ToLower() switch
        {
            "get" => GetProperties(path, sessionId),
            "set" => SetProperties(path, sessionId, outputPath, title, subject, author, keywords, comments, category,
                company, manager, customProperties),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets presentation properties using IPresentationInfo for efficiency.
    ///     This method reads properties without loading the entire presentation.
    /// </summary>
    /// <param name="path">The presentation file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <returns>A JSON string containing the document properties.</returns>
    /// <exception cref="ArgumentException">Thrown when neither sessionId nor path is provided.</exception>
    private string GetProperties(string? path, string? sessionId)
    {
        if (!string.IsNullOrEmpty(sessionId))
        {
            using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);
            var props = ctx.Document.DocumentProperties;

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

            return JsonSerializer.Serialize(result, JsonOptions);
        }
        else
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Either sessionId or path must be provided");

            SecurityHelper.ValidateFilePath(path, "path", true);

            var info = PresentationFactory.Instance.GetPresentationInfo(path);
            var props = info.ReadDocumentProperties();

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

            return JsonSerializer.Serialize(result, JsonOptions);
        }
    }

    /// <summary>
    ///     Sets presentation properties.
    /// </summary>
    /// <param name="path">The presentation file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="title">The document title.</param>
    /// <param name="subject">The document subject.</param>
    /// <param name="author">The document author.</param>
    /// <param name="keywords">The document keywords.</param>
    /// <param name="comments">The document comments.</param>
    /// <param name="category">The document category.</param>
    /// <param name="company">The document company.</param>
    /// <param name="manager">The document manager.</param>
    /// <param name="customProperties">The custom properties as key-value pairs.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string SetProperties(string? path, string? sessionId, string? outputPath, string? title, string? subject,
        string? author, string? keywords, string? comments, string? category, string? company, string? manager,
        Dictionary<string, object>? customProperties)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        var props = ctx.Document.DocumentProperties;
        List<string> changes = [];

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

        if (customProperties != null)
        {
            foreach (var kvp in customProperties)
                props[kvp.Key] = ConvertToPropertyValue(kvp.Value);
            changes.Add("CustomProperties");
        }

        ctx.Save(outputPath);

        var result = $"Document properties updated: {string.Join(", ", changes)}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Converts a dictionary value to proper property value type.
    /// </summary>
    /// <param name="value">The value to convert.</param>
    /// <returns>The converted property value.</returns>
    private static object ConvertToPropertyValue(object value)
    {
        if (value is JsonElement element)
            return element.ValueKind switch
            {
                JsonValueKind.String => TryParseDateTime(element.GetString()!, out var dt) ? dt : element.GetString()!,
                JsonValueKind.Number => element.TryGetInt32(out var intVal) ? intVal : element.GetDouble(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                _ => element.ToString()
            };

        return value;
    }

    /// <summary>
    ///     Attempts to parse a string as a DateTime value.
    /// </summary>
    /// <param name="value">The string value to parse.</param>
    /// <param name="result">When successful, contains the parsed DateTime value.</param>
    /// <returns>True if parsing succeeded; otherwise, false.</returns>
    private static bool TryParseDateTime(string value, out DateTime result)
    {
        return DateTime.TryParse(value, out result);
    }
}