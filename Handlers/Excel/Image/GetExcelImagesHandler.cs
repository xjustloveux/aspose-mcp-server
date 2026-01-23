using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Image;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Handler for getting images from Excel worksheets.
/// </summary>
[ResultType(typeof(GetImagesExcelResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var pictures = worksheet.Pictures;

        if (pictures.Count == 0)
            return new GetImagesExcelResult
            {
                Count = 0,
                WorksheetName = worksheet.Name,
                Items = Array.Empty<ExcelImageInfo>(),
                Message = "No images found"
            };

        List<ExcelImageInfo> imageList = [];
        for (var i = 0; i < pictures.Count; i++)
        {
            var picture = pictures[i];
            var upperLeftCell = CellsHelper.CellIndexToName(picture.UpperLeftRow, picture.UpperLeftColumn);
            var lowerRightCell = CellsHelper.CellIndexToName(picture.LowerRightRow, picture.LowerRightColumn);

            imageList.Add(new ExcelImageInfo
            {
                Index = i,
                Name = picture.Name,
                AlternativeText = picture.AlternativeText,
                ImageType = picture.ImageType.ToString(),
                Location = new ExcelImageLocation
                {
                    UpperLeftCell = upperLeftCell,
                    LowerRightCell = lowerRightCell,
                    UpperLeftRow = picture.UpperLeftRow,
                    UpperLeftColumn = picture.UpperLeftColumn,
                    LowerRightRow = picture.LowerRightRow,
                    LowerRightColumn = picture.LowerRightColumn
                },
                Width = picture.Width,
                Height = picture.Height,
                IsLockAspectRatio = picture.IsLockAspectRatio
            });
        }

        return new GetImagesExcelResult
        {
            Count = pictures.Count,
            WorksheetName = worksheet.Name,
            Items = imageList
        };
    }

    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        return new GetParameters(sheetIndex);
    }

    private sealed record GetParameters(int SheetIndex);
}
