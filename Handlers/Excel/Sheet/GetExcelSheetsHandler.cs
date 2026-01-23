using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Excel.Sheet;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for getting worksheet information from Excel workbooks.
/// </summary>
[ResultType(typeof(GetSheetsResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        _ = parameters;

        var workbook = context.Document;

        var workbookName = context.SourcePath != null ? Path.GetFileName(context.SourcePath) : "session";

        if (workbook.Worksheets.Count == 0)
            return new GetSheetsResult
            {
                Count = 0,
                WorkbookName = workbookName,
                Items = [],
                Message = "No worksheets found"
            };

        List<ExcelSheetInfo> sheetList = [];
        for (var i = 0; i < workbook.Worksheets.Count; i++)
        {
            var worksheet = workbook.Worksheets[i];
            sheetList.Add(new ExcelSheetInfo
            {
                Index = i,
                Name = worksheet.Name,
                Visibility = worksheet.IsVisible ? "Visible" : "Hidden"
            });
        }

        return new GetSheetsResult
        {
            Count = workbook.Worksheets.Count,
            WorkbookName = workbookName,
            Items = sheetList
        };
    }
}
