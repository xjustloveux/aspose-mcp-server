using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Excel.Properties;

namespace AsposeMcpServer.Handlers.Excel.Properties;

/// <summary>
///     Handler for getting workbook properties from Excel files.
/// </summary>
[ResultType(typeof(GetWorkbookPropertiesResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var workbook = context.Document;
        var props = workbook.BuiltInDocumentProperties;
        var customProps = workbook.CustomDocumentProperties;

        List<CustomPropertyInfo> customPropsList = [];
        if (customProps.Count > 0)
            foreach (var prop in customProps)
                customPropsList.Add(new CustomPropertyInfo
                {
                    Name = prop.Name,
                    Value = prop.Value?.ToString(),
                    Type = prop.Type.ToString()
                });

        return new GetWorkbookPropertiesResult
        {
            Title = props.Title,
            Subject = props.Subject,
            Author = props.Author,
            Keywords = props.Keywords,
            Comments = props.Comments,
            Category = props.Category,
            Company = props.Company,
            Manager = props.Manager,
            Created = props.CreatedTime.ToString("o"),
            Modified = props.LastSavedTime.ToString("o"),
            LastSavedBy = props.LastSavedBy,
            Revision = props.RevisionNumber,
            TotalSheets = workbook.Worksheets.Count,
            CustomProperties = customPropsList
        };
    }
}
