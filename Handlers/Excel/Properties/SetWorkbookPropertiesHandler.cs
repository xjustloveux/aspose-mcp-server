using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Properties;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Properties;

/// <summary>
///     Handler for setting workbook properties in Excel files.
/// </summary>
public class SetWorkbookPropertiesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_workbook_properties";

    /// <summary>
    ///     Sets workbook properties including built-in and custom properties.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: title, subject, author, keywords, comments, category, company, manager, customProperties
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
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

        var workbook = context.Document;
        var props = workbook.BuiltInDocumentProperties;

        if (!string.IsNullOrEmpty(title)) props.Title = title;
        if (!string.IsNullOrEmpty(subject)) props.Subject = subject;
        if (!string.IsNullOrEmpty(author)) props.Author = author;
        if (!string.IsNullOrEmpty(keywords)) props.Keywords = keywords;
        if (!string.IsNullOrEmpty(comments)) props.Comments = comments;
        if (!string.IsNullOrEmpty(category)) props.Category = category;
        if (!string.IsNullOrEmpty(company)) props.Company = company;
        if (!string.IsNullOrEmpty(manager)) props.Manager = manager;

        if (!string.IsNullOrEmpty(customPropertiesJson))
            try
            {
                var customProps = JsonNode.Parse(customPropertiesJson)?.AsObject();
                if (customProps != null)
                    foreach (var kvp in customProps)
                    {
                        var value = kvp.Value?.GetValue<string>() ?? "";
                        var existingProp = FindCustomProperty(workbook.CustomDocumentProperties, kvp.Key);
                        if (existingProp != null)
                            workbook.CustomDocumentProperties.Remove(kvp.Key);
                        workbook.CustomDocumentProperties.Add(kvp.Key, value);
                    }
            }
            catch (JsonException ex)
            {
                throw new ArgumentException($"Invalid JSON format for customProperties: {ex.Message}");
            }

        MarkModified(context);
        return Success("Workbook properties updated successfully.");
    }

    /// <summary>
    ///     Finds a custom property by name (case-insensitive).
    /// </summary>
    /// <param name="customProperties">The custom properties collection to search.</param>
    /// <param name="name">The property name to find.</param>
    /// <returns>The found property or null if not found.</returns>
    private static DocumentProperty? FindCustomProperty(CustomDocumentPropertyCollection customProperties, string name)
    {
        foreach (var prop in customProperties)
            if (string.Equals(prop.Name, name, StringComparison.OrdinalIgnoreCase))
                return prop;
        return null;
    }
}
