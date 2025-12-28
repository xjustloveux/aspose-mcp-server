using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word document properties (get, set)
///     Merges: WordGetDocumentPropertiesTool, WordSetDocumentPropertiesTool, WordSetPropertiesTool
/// </summary>
public class WordPropertiesTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Get or set Word document properties (metadata). Supports 2 operations: get, set.

Usage examples:
- Get properties: word_properties(operation='get', path='doc.docx')
- Set properties: word_properties(operation='set', path='doc.docx', title='Title', author='Author', subject='Subject')

Notes:
- The 'set' operation is for content metadata (title, author, subject, etc.), not for statistics (word count, page count)
- Statistics like word count and page count are automatically calculated by Word and cannot be manually set
- Custom properties support multiple types: string, number (integer/double), boolean, and datetime (ISO 8601 format)";

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
                description =
                    "Custom properties as key-value pairs (optional, for set operation). Supports string, number (integer/double), boolean, and datetime (ISO 8601 format, e.g., '2024-01-15T10:30:00'). If a property key already exists, it will be updated."
            }
        },
        required = new[] { "operation", "path" }
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
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "get" => await GetPropertiesAsync(path),
            "set" => await SetPropertiesAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets document properties including built-in and custom properties
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <returns>JSON string with document properties for easy parsing by LLM and tool chains</returns>
    private Task<string> GetPropertiesAsync(string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);

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
        });
    }

    /// <summary>
    ///     Converts a property value to JsonNode for safe serialization
    /// </summary>
    /// <param name="value">Property value object</param>
    /// <returns>JsonNode representation of the value</returns>
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
    ///     Sets document properties including built-in and custom properties
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing various property values</param>
    /// <returns>Success message with output path</returns>
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

            doc.Save(outputPath);
            return $"Document properties updated: {outputPath}";
        });
    }

    /// <summary>
    ///     Adds a custom property with the appropriate type based on JSON value
    /// </summary>
    /// <param name="doc">Word document</param>
    /// <param name="key">Property key name</param>
    /// <param name="jsonValue">JSON value node</param>
    private void AddCustomPropertyWithType(Document doc, string key, JsonNode? jsonValue)
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