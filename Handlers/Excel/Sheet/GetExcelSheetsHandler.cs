using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for getting worksheet information from Excel workbooks.
/// </summary>
public class GetExcelSheetsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets information about all worksheets in the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sourcePath (for display in result)
    /// </param>
    /// <returns>JSON string containing information about all worksheets.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        _ = parameters;

        var workbook = context.Document;

        if (workbook.Worksheets.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                workbookName = context.SourcePath != null ? Path.GetFileName(context.SourcePath) : "session",
                items = Array.Empty<object>(),
                message = "No worksheets found"
            };
            return JsonResult(emptyResult);
        }

        List<object> sheetList = [];
        for (var i = 0; i < workbook.Worksheets.Count; i++)
        {
            var worksheet = workbook.Worksheets[i];
            sheetList.Add(new
            {
                index = i,
                name = worksheet.Name,
                visibility = worksheet.IsVisible ? "Visible" : "Hidden"
            });
        }

        var result = new
        {
            count = workbook.Worksheets.Count,
            workbookName = context.SourcePath != null ? Path.GetFileName(context.SourcePath) : "session",
            items = sheetList
        };

        return JsonResult(result);
    }
}
