using System.Globalization;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Properties;

/// <summary>
///     Handler for setting Word document properties.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object
        Execute(OperationContext<Document> context,
            OperationParameters parameters)
    {
        var setParams = ExtractSetPropertiesParameters(parameters);

        var doc = context.Document;
        var props = doc.BuiltInDocumentProperties;

        if (!string.IsNullOrEmpty(setParams.Title)) props.Title = setParams.Title;
        if (!string.IsNullOrEmpty(setParams.Subject)) props.Subject = setParams.Subject;
        if (!string.IsNullOrEmpty(setParams.Author)) props.Author = setParams.Author;
        if (!string.IsNullOrEmpty(setParams.Keywords)) props.Keywords = setParams.Keywords;
        if (!string.IsNullOrEmpty(setParams.Comments)) props.Comments = setParams.Comments;
        if (!string.IsNullOrEmpty(setParams.Category)) props.Category = setParams.Category;
        if (!string.IsNullOrEmpty(setParams.Company)) props.Company = setParams.Company;
        if (!string.IsNullOrEmpty(setParams.Manager)) props.Manager = setParams.Manager;

        if (!string.IsNullOrEmpty(setParams.CustomPropertiesJson))
        {
            var customProps = JsonNode.Parse(setParams.CustomPropertiesJson)?.AsObject();
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

        return new SuccessResult { Message = "Document properties updated" };
    }

    /// <summary>
    ///     Adds a custom property with the appropriate type based on JSON value.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="key">The property key.</param>
    /// <param name="jsonValue">The JSON value.</param>
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
                if (DateTime.TryParse(strVal, CultureInfo.InvariantCulture, out var dateVal))
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

    /// <summary>
    ///     Extracts set properties parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set properties parameters.</returns>
    private static SetPropertiesParameters ExtractSetPropertiesParameters(OperationParameters parameters)
    {
        return new SetPropertiesParameters(
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<string?>("subject"),
            parameters.GetOptional<string?>("author"),
            parameters.GetOptional<string?>("keywords"),
            parameters.GetOptional<string?>("comments"),
            parameters.GetOptional<string?>("category"),
            parameters.GetOptional<string?>("company"),
            parameters.GetOptional<string?>("manager"),
            parameters.GetOptional<string?>("customProperties")
        );
    }

    /// <summary>
    ///     Record to hold set properties parameters.
    /// </summary>
    /// <param name="Title">The document title.</param>
    /// <param name="Subject">The document subject.</param>
    /// <param name="Author">The document author.</param>
    /// <param name="Keywords">The document keywords.</param>
    /// <param name="Comments">The document comments.</param>
    /// <param name="Category">The document category.</param>
    /// <param name="Company">The document company.</param>
    /// <param name="Manager">The document manager.</param>
    /// <param name="CustomPropertiesJson">The custom properties as JSON string.</param>
    private sealed record SetPropertiesParameters(
        string? Title,
        string? Subject,
        string? Author,
        string? Keywords,
        string? Comments,
        string? Category,
        string? Company,
        string? Manager,
        string? CustomPropertiesJson);
}
