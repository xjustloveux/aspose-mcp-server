using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word document properties (get, set)
///     Merges: WordGetDocumentPropertiesTool, WordSetDocumentPropertiesTool, WordSetPropertiesTool
/// </summary>
[McpServerToolType]
public class WordPropertiesTool
{
    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordPropertiesTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    public WordPropertiesTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "word_properties")]
    [Description(@"Get or set Word document properties (metadata). Supports 2 operations: get, set.

Usage examples:
- Get properties: word_properties(operation='get', path='doc.docx')
- Set properties: word_properties(operation='set', path='doc.docx', title='Title', author='Author', subject='Subject')

Notes:
- The 'set' operation is for content metadata (title, author, subject, etc.), not for statistics (word count, page count)
- Statistics like word count and page count are automatically calculated by Word and cannot be manually set
- Custom properties support multiple types: string, number (integer/double), boolean, and datetime (ISO 8601 format)")]
    public string Execute(
        [Description("Operation: get, set")] string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (if not provided, overwrites input, for set operation)")]
        string? outputPath = null,
        [Description("Document title (optional, for set operation)")]
        string? title = null,
        [Description("Document subject (optional, for set operation)")]
        string? subject = null,
        [Description("Document author (optional, for set operation)")]
        string? author = null,
        [Description("Keywords (optional, for set operation)")]
        string? keywords = null,
        [Description("Comments (optional, for set operation)")]
        string? comments = null,
        [Description("Document category (optional, for set operation)")]
        string? category = null,
        [Description("Company name (optional, for set operation)")]
        string? company = null,
        [Description("Manager name (optional, for set operation)")]
        string? manager = null,
        [Description(
            "Custom properties as JSON string (optional, for set operation). Supports string, number (integer/double), boolean, and datetime (ISO 8601 format).")]
        string? customProperties = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "get" => GetProperties(ctx),
            "set" => SetProperties(ctx, outputPath, title, subject, author, keywords, comments, category, company,
                manager, customProperties),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets document properties including built-in and custom properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing built-in properties, statistics, and custom properties.</returns>
    private static string GetProperties(DocumentContext<Document> ctx)
    {
        var doc = ctx.Document;

        // Update word count and statistics to ensure accurate metadata
        doc.UpdateWordCount();

        var props = doc.BuiltInDocumentProperties;
        var customProps = doc.CustomDocumentProperties;

        var result = new JsonObject
        {
            ["builtInProperties"] = new JsonObject
            {
                ["title"] = props.Title,
                ["subject"] = props.Subject,
                ["author"] = props.Author,
                ["keywords"] = props.Keywords,
                ["comments"] = props.Comments,
                ["category"] = props.Category,
                ["company"] = props.Company,
                ["manager"] = props.Manager,
                ["createdTime"] = props.CreatedTime.ToString("O"),
                ["lastSavedTime"] = props.LastSavedTime.ToString("O"),
                ["lastSavedBy"] = props.LastSavedBy,
                ["revisionNumber"] = props.RevisionNumber
            },
            ["statistics"] = new JsonObject
            {
                ["wordCount"] = props.Words,
                ["characterCount"] = props.Characters,
                ["pageCount"] = props.Pages,
                ["paragraphCount"] = props.Paragraphs,
                ["lineCount"] = props.Lines
            }
        };

        if (customProps.Count > 0)
        {
            var customPropsJson = new JsonObject();
            foreach (var prop in customProps)
                customPropsJson[prop.Name] = new JsonObject
                {
                    ["value"] = ConvertPropertyValueToJsonNode(prop.Value),
                    ["type"] = prop.Type.ToString()
                };
            result["customProperties"] = customPropsJson;
        }

        return result.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Converts a property value to JsonNode for safe serialization.
    /// </summary>
    /// <param name="value">The property value to convert.</param>
    /// <returns>A JsonNode representation of the value, or null if the value is null.</returns>
    private static JsonNode? ConvertPropertyValueToJsonNode(object? value)
    {
        return value switch
        {
            null => null,
            bool b => JsonValue.Create(b),
            int i => JsonValue.Create(i),
            double d => JsonValue.Create(d),
            DateTime dt => JsonValue.Create(dt.ToString("O")),
            _ => JsonValue.Create(value.ToString())
        };
    }

    /// <summary>
    ///     Sets document properties including built-in and custom properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="title">The document title.</param>
    /// <param name="subject">The document subject.</param>
    /// <param name="author">The document author.</param>
    /// <param name="keywords">The document keywords.</param>
    /// <param name="comments">The document comments.</param>
    /// <param name="category">The document category.</param>
    /// <param name="company">The company name.</param>
    /// <param name="manager">The manager name.</param>
    /// <param name="customPropertiesJson">Custom properties as JSON string.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetProperties(
        DocumentContext<Document> ctx,
        string? outputPath,
        string? title,
        string? subject,
        string? author,
        string? keywords,
        string? comments,
        string? category,
        string? company,
        string? manager,
        string? customPropertiesJson)
    {
        var doc = ctx.Document;
        var props = doc.BuiltInDocumentProperties;

        if (!string.IsNullOrEmpty(title)) props.Title = title;
        if (!string.IsNullOrEmpty(subject)) props.Subject = subject;
        if (!string.IsNullOrEmpty(author)) props.Author = author;
        if (!string.IsNullOrEmpty(keywords)) props.Keywords = keywords;
        if (!string.IsNullOrEmpty(comments)) props.Comments = comments;
        if (!string.IsNullOrEmpty(category)) props.Category = category;
        if (!string.IsNullOrEmpty(company)) props.Company = company;
        if (!string.IsNullOrEmpty(manager)) props.Manager = manager;

        if (!string.IsNullOrEmpty(customPropertiesJson))
        {
            var customProps = JsonNode.Parse(customPropertiesJson)?.AsObject();
            if (customProps != null)
                foreach (var kvp in customProps)
                {
                    var key = kvp.Key;
                    var jsonValue = kvp.Value;

                    // Remove existing property if it exists to allow update
                    if (doc.CustomDocumentProperties[key] != null)
                        doc.CustomDocumentProperties.Remove(key);

                    // Add property with appropriate type based on JSON value
                    AddCustomPropertyWithType(doc, key, jsonValue);
                }
        }

        ctx.Save(outputPath);

        var result = "Document properties updated\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Adds a custom property with the appropriate type based on JSON value.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="key">The property key name.</param>
    /// <param name="jsonValue">The JSON value to add as a custom property.</param>
    private static void AddCustomPropertyWithType(Document doc, string key, JsonNode? jsonValue)
    {
        if (jsonValue == null)
        {
            doc.CustomDocumentProperties.Add(key, string.Empty);
            return;
        }

        // Try to determine the type from the JSON value
        if (jsonValue is JsonValue jv)
        {
            // Try boolean first
            if (jv.TryGetValue<bool>(out var boolVal))
            {
                doc.CustomDocumentProperties.Add(key, boolVal);
                return;
            }

            // Try integer
            if (jv.TryGetValue<int>(out var intVal))
            {
                doc.CustomDocumentProperties.Add(key, intVal);
                return;
            }

            // Try double
            if (jv.TryGetValue<double>(out var doubleVal))
            {
                doc.CustomDocumentProperties.Add(key, doubleVal);
                return;
            }

            // Try datetime (ISO 8601 format)
            if (jv.TryGetValue<string>(out var strVal) && !string.IsNullOrEmpty(strVal))
            {
                if (DateTime.TryParse(strVal, out var dateVal))
                {
                    doc.CustomDocumentProperties.Add(key, dateVal);
                    return;
                }

                // Default to string
                doc.CustomDocumentProperties.Add(key, strVal);
                return;
            }
        }

        // Fallback to string representation
        doc.CustomDocumentProperties.Add(key, jsonValue.ToString());
    }
}