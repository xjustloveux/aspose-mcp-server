using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Properties;

/// <summary>
///     Handler for getting workbook properties from Excel files.
/// </summary>
public class GetWorkbookPropertiesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_workbook_properties";

    /// <summary>
    ///     Gets workbook properties including built-in and custom properties.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">No additional parameters required.</param>
    /// <returns>JSON result with workbook properties.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var workbook = context.Document;
        var props = workbook.BuiltInDocumentProperties;
        var customProps = workbook.CustomDocumentProperties;

        List<object> customPropsList = [];
        if (customProps.Count > 0)
            foreach (var prop in customProps)
                customPropsList.Add(new
                    { name = prop.Name, value = prop.Value?.ToString(), type = prop.Type.ToString() });

        var result = new
        {
            title = props.Title,
            subject = props.Subject,
            author = props.Author,
            keywords = props.Keywords,
            comments = props.Comments,
            category = props.Category,
            company = props.Company,
            manager = props.Manager,
            created = props.CreatedTime.ToString("o"),
            modified = props.LastSavedTime.ToString("o"),
            lastSavedBy = props.LastSavedBy,
            revision = props.RevisionNumber,
            totalSheets = workbook.Worksheets.Count,
            customProperties = customPropsList
        };

        return JsonResult(result);
    }
}
