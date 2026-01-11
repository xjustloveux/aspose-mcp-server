using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Properties;

/// <summary>
///     Handler for setting PowerPoint presentation properties.
/// </summary>
public class SetPptPropertiesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set";

    /// <summary>
    ///     Sets presentation properties.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: title, subject, author, keywords, comments, category, company, manager, customProperties
    /// </param>
    /// <returns>Success message with updated properties list.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var title = parameters.GetOptional<string?>("title");
        var subject = parameters.GetOptional<string?>("subject");
        var author = parameters.GetOptional<string?>("author");
        var keywords = parameters.GetOptional<string?>("keywords");
        var comments = parameters.GetOptional<string?>("comments");
        var category = parameters.GetOptional<string?>("category");
        var company = parameters.GetOptional<string?>("company");
        var manager = parameters.GetOptional<string?>("manager");
        var customProperties = parameters.GetOptional<Dictionary<string, object>?>("customProperties");

        var presentation = context.Document;
        var props = presentation.DocumentProperties;
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

        MarkModified(context);

        return Success($"Document properties updated: {string.Join(", ", changes)}.");
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
