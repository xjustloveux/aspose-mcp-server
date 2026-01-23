using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Properties;

namespace AsposeMcpServer.Handlers.Excel.Properties;

/// <summary>
///     Handler for getting worksheet properties from Excel files.
/// </summary>
[ResultType(typeof(GetSheetPropertiesResult))]
public class GetSheetPropertiesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_sheet_properties";

    /// <summary>
    ///     Gets worksheet properties including data counts, protection status, and print settings.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sheetIndex (0-based)
    /// </param>
    /// <returns>JSON result with worksheet properties.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetSheetPropertiesParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var pageSetup = worksheet.PageSetup;

        return new GetSheetPropertiesResult
        {
            Name = worksheet.Name,
            Index = getParams.SheetIndex,
            IsVisible = worksheet.IsVisible,
            TabColor = worksheet.TabColor.ToString(),
            IsSelected = workbook.Worksheets.ActiveSheetIndex == getParams.SheetIndex,
            DataRowCount = worksheet.Cells.MaxDataRow + 1,
            DataColumnCount = worksheet.Cells.MaxDataColumn + 1,
            IsProtected = worksheet.Protection.IsProtectedWithPassword,
            CommentsCount = worksheet.Comments.Count,
            ChartsCount = worksheet.Charts.Count,
            PicturesCount = worksheet.Pictures.Count,
            HyperlinksCount = worksheet.Hyperlinks.Count,
            PrintSettings = new PrintSettingsInfo
            {
                PrintArea = pageSetup.PrintArea,
                PrintTitleRows = pageSetup.PrintTitleRows,
                PrintTitleColumns = pageSetup.PrintTitleColumns,
                Orientation = pageSetup.Orientation.ToString(),
                PaperSize = pageSetup.PaperSize.ToString(),
                FitToPagesWide = pageSetup.FitToPagesWide,
                FitToPagesTall = pageSetup.FitToPagesTall
            }
        };
    }

    /// <summary>
    ///     Extracts get sheet properties parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get sheet properties parameters.</returns>
    private static GetSheetPropertiesParameters ExtractGetSheetPropertiesParameters(OperationParameters parameters)
    {
        return new GetSheetPropertiesParameters(
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    /// <summary>
    ///     Record to hold get sheet properties parameters.
    /// </summary>
    private sealed record GetSheetPropertiesParameters(int SheetIndex);
}
