using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Properties;

/// <summary>
///     Handler for getting Word document properties.
/// </summary>
public class GetWordPropertiesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets document properties including built-in and custom properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>A JSON string containing built-in properties, statistics, and custom properties.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;

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
}
