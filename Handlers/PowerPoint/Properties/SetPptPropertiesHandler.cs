using System.Globalization;
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
        var p = ExtractSetPptPropertiesParameters(parameters);
        var presentation = context.Document;
        var props = presentation.DocumentProperties;
        List<string> changes = [];

        if (!string.IsNullOrEmpty(p.Title))
        {
            props.Title = p.Title;
            changes.Add("Title");
        }

        if (!string.IsNullOrEmpty(p.Subject))
        {
            props.Subject = p.Subject;
            changes.Add("Subject");
        }

        if (!string.IsNullOrEmpty(p.Author))
        {
            props.Author = p.Author;
            changes.Add("Author");
        }

        if (!string.IsNullOrEmpty(p.Keywords))
        {
            props.Keywords = p.Keywords;
            changes.Add("Keywords");
        }

        if (!string.IsNullOrEmpty(p.Comments))
        {
            props.Comments = p.Comments;
            changes.Add("Comments");
        }

        if (!string.IsNullOrEmpty(p.Category))
        {
            props.Category = p.Category;
            changes.Add("Category");
        }

        if (!string.IsNullOrEmpty(p.Company))
        {
            props.Company = p.Company;
            changes.Add("Company");
        }

        if (!string.IsNullOrEmpty(p.Manager))
        {
            props.Manager = p.Manager;
            changes.Add("Manager");
        }

        if (p.CustomProperties != null)
        {
            foreach (var kvp in p.CustomProperties)
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
        return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
    }

    private static SetPptPropertiesParameters ExtractSetPptPropertiesParameters(OperationParameters parameters)
    {
        return new SetPptPropertiesParameters(
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<string?>("subject"),
            parameters.GetOptional<string?>("author"),
            parameters.GetOptional<string?>("keywords"),
            parameters.GetOptional<string?>("comments"),
            parameters.GetOptional<string?>("category"),
            parameters.GetOptional<string?>("company"),
            parameters.GetOptional<string?>("manager"),
            parameters.GetOptional<Dictionary<string, object>?>("customProperties"));
    }

    private sealed record SetPptPropertiesParameters(
        string? Title,
        string? Subject,
        string? Author,
        string? Keywords,
        string? Comments,
        string? Category,
        string? Company,
        string? Manager,
        Dictionary<string, object>? CustomProperties);
}
