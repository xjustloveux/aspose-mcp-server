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
        var setParams = ExtractSetWorkbookPropertiesParameters(parameters);

        var workbook = context.Document;
        var props = workbook.BuiltInDocumentProperties;

        SetBuiltInProperties(props, setParams);
        SetCustomProperties(workbook, setParams.CustomProperties);

        MarkModified(context);
        return Success("Workbook properties updated successfully.");
    }

    /// <summary>
    ///     Sets the built-in document properties from the parameters.
    /// </summary>
    /// <param name="props">The built-in document properties collection.</param>
    /// <param name="setParams">The set workbook properties parameters.</param>
    private static void SetBuiltInProperties(BuiltInDocumentPropertyCollection props,
        SetWorkbookPropertiesParameters setParams)
    {
        if (!string.IsNullOrEmpty(setParams.Title)) props.Title = setParams.Title;
        if (!string.IsNullOrEmpty(setParams.Subject)) props.Subject = setParams.Subject;
        if (!string.IsNullOrEmpty(setParams.Author)) props.Author = setParams.Author;
        if (!string.IsNullOrEmpty(setParams.Keywords)) props.Keywords = setParams.Keywords;
        if (!string.IsNullOrEmpty(setParams.Comments)) props.Comments = setParams.Comments;
        if (!string.IsNullOrEmpty(setParams.Category)) props.Category = setParams.Category;
        if (!string.IsNullOrEmpty(setParams.Company)) props.Company = setParams.Company;
        if (!string.IsNullOrEmpty(setParams.Manager)) props.Manager = setParams.Manager;
    }

    /// <summary>
    ///     Sets the custom document properties from a JSON string.
    /// </summary>
    /// <param name="workbook">The workbook to update.</param>
    /// <param name="customPropertiesJson">The JSON string containing custom properties.</param>
    /// <exception cref="ArgumentException">Thrown when the JSON format is invalid.</exception>
    private static void SetCustomProperties(Workbook workbook, string? customPropertiesJson)
    {
        if (string.IsNullOrEmpty(customPropertiesJson)) return;

        try
        {
            var customProps = JsonNode.Parse(customPropertiesJson)?.AsObject();
            if (customProps == null) return;

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
    }

    /// <summary>
    ///     Finds a custom property by name (case-insensitive).
    /// </summary>
    /// <param name="customProperties">The custom properties collection to search.</param>
    /// <param name="name">The property name to find.</param>
    /// <returns>The found property or null if not found.</returns>
    private static DocumentProperty? FindCustomProperty(CustomDocumentPropertyCollection customProperties, string name)
    {
        return customProperties
            .FirstOrDefault(prop => string.Equals(prop.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    ///     Extracts set workbook properties parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set workbook properties parameters.</returns>
    private static SetWorkbookPropertiesParameters ExtractSetWorkbookPropertiesParameters(
        OperationParameters parameters)
    {
        return new SetWorkbookPropertiesParameters(
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
    ///     Record to hold set workbook properties parameters.
    /// </summary>
    private sealed record SetWorkbookPropertiesParameters(
        string? Title,
        string? Subject,
        string? Author,
        string? Keywords,
        string? Comments,
        string? Category,
        string? Company,
        string? Manager,
        string? CustomProperties);
}
