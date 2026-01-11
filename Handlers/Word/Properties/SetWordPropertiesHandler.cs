using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Properties;

/// <summary>
///     Handler for setting Word document properties.
/// </summary>
public class SetWordPropertiesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set";

    /// <summary>
    ///     Sets document properties including built-in and custom properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: title, subject, author, keywords, comments, category, company, manager, customProperties
    /// </param>
    /// <returns>Success message indicating properties were updated.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var title = parameters.GetOptional<string?>("title");
        var subject = parameters.GetOptional<string?>("subject");
        var author = parameters.GetOptional<string?>("author");
        var keywords = parameters.GetOptional<string?>("keywords");
        var comments = parameters.GetOptional<string?>("comments");
        var category = parameters.GetOptional<string?>("category");
        var company = parameters.GetOptional<string?>("company");
        var manager = parameters.GetOptional<string?>("manager");
        var customPropertiesJson = parameters.GetOptional<string?>("customProperties");

        var doc = context.Document;
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

                    if (doc.CustomDocumentProperties[key] != null)
                        doc.CustomDocumentProperties.Remove(key);

                    AddCustomPropertyWithType(doc, key, jsonValue);
                }
        }

        MarkModified(context);

        return Success("Document properties updated");
    }

    /// <summary>
    ///     Adds a custom property with the appropriate type based on JSON value.
    /// </summary>
    private static void AddCustomPropertyWithType(Document doc, string key, JsonNode? jsonValue)
    {
        if (jsonValue == null)
        {
            doc.CustomDocumentProperties.Add(key, string.Empty);
            return;
        }

        if (jsonValue is JsonValue jv)
        {
            if (jv.TryGetValue<bool>(out var boolVal))
            {
                doc.CustomDocumentProperties.Add(key, boolVal);
                return;
            }

            if (jv.TryGetValue<int>(out var intVal))
            {
                doc.CustomDocumentProperties.Add(key, intVal);
                return;
            }

            if (jv.TryGetValue<double>(out var doubleVal))
            {
                doc.CustomDocumentProperties.Add(key, doubleVal);
                return;
            }

            if (jv.TryGetValue<string>(out var strVal) && !string.IsNullOrEmpty(strVal))
            {
                if (DateTime.TryParse(strVal, out var dateVal))
                {
                    doc.CustomDocumentProperties.Add(key, dateVal);
                    return;
                }

                doc.CustomDocumentProperties.Add(key, strVal);
                return;
            }
        }

        doc.CustomDocumentProperties.Add(key, jsonValue.ToString());
    }
}
