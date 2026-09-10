using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Properties;

/// <summary>
///     Handler for setting PowerPoint presentation properties.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSetPptPropertiesParameters(parameters);
        var presentation = context.Document;
        var props = presentation.DocumentProperties;
        List<string> changes = [];

        // null means the caller did not supply the field; an empty string means clear it.
        // Testing IsNullOrEmpty collapsed the two, so a field could never be cleared.
        if (p.Title != null)
        {
            props.Title = p.Title;
            changes.Add("Title");
        }

        if (p.Subject != null)
        {
            props.Subject = p.Subject;
            changes.Add("Subject");
        }

        if (p.Author != null)
        {
            props.Author = p.Author;
            changes.Add("Author");
        }

        if (p.Keywords != null)
        {
            props.Keywords = p.Keywords;
            changes.Add("Keywords");
        }

        if (p.Comments != null)
        {
            props.Comments = p.Comments;
            changes.Add("Comments");
        }

        if (p.Category != null)
        {
            props.Category = p.Category;
            changes.Add("Category");
        }

        if (p.Company != null)
        {
            props.Company = p.Company;
            changes.Add("Company");
        }

        if (p.Manager != null)
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

        return new SuccessResult { Message = $"Document properties updated: {string.Join(", ", changes)}." };
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
        return IsoDateTimeHelper.TryParse(value, out result);
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
