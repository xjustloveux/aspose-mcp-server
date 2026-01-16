using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Handler for getting images from Excel worksheets.
/// </summary>
public class GetExcelImagesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all images from the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>JSON result with image information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var pictures = worksheet.Pictures;

        if (pictures.Count == 0)
            return JsonResult(new
            {
                count = 0,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No images found"
            });

        List<object> imageList = [];
        for (var i = 0; i < pictures.Count; i++)
        {
            var picture = pictures[i];
            var upperLeftCell = CellsHelper.CellIndexToName(picture.UpperLeftRow, picture.UpperLeftColumn);
            var lowerRightCell = CellsHelper.CellIndexToName(picture.LowerRightRow, picture.LowerRightColumn);

            imageList.Add(new
            {
                index = i,
                name = picture.Name,
                alternativeText = picture.AlternativeText,
                imageType = picture.ImageType.ToString(),
                location = new
                {
                    upperLeftCell,
                    lowerRightCell,
                    upperLeftRow = picture.UpperLeftRow,
                    upperLeftColumn = picture.UpperLeftColumn,
                    lowerRightRow = picture.LowerRightRow,
                    lowerRightColumn = picture.LowerRightColumn
                },
                width = picture.Width,
                height = picture.Height,
                isLockAspectRatio = picture.IsLockAspectRatio
            });
        }

        return JsonResult(new
        {
            count = pictures.Count,
            worksheetName = worksheet.Name,
            items = imageList
        });
    }

    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        return new GetParameters(sheetIndex);
    }

    private sealed record GetParameters(int SheetIndex);
}
