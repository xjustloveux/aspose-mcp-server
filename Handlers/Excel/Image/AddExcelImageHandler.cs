using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Handler for adding images to Excel worksheets.
/// </summary>
public class AddExcelImageHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds an image to the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: imagePath, cell
    ///     Optional: sheetIndex (default: 0), width, height, keepAspectRatio (default: true)
    /// </param>
    /// <returns>Success message with image details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var imagePath = parameters.GetRequired<string>("imagePath");
        var cell = parameters.GetRequired<string>("cell");
        var width = parameters.GetOptional<int?>("width");
        var height = parameters.GetOptional<int?>("height");
        var keepAspectRatio = parameters.GetOptional("keepAspectRatio", true);

        SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        ExcelImageHelper.ValidateImageFormat(imagePath);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        var pictureIndex = worksheet.Pictures.Add(cellObj.Row, cellObj.Column, imagePath);
        var picture = worksheet.Pictures[pictureIndex];

        if (width.HasValue || height.HasValue)
        {
            picture.IsLockAspectRatio = keepAspectRatio;
            if (width.HasValue) picture.Width = width.Value;
            if (height.HasValue) picture.Height = height.Value;
        }

        MarkModified(context);

        return Success($"Image added to cell {cell} (size: {picture.Width}x{picture.Height}).");
    }
}
