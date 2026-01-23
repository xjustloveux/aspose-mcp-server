using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Handler for adding images to Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var addParams = ExtractAddParameters(parameters);

        SecurityHelper.ValidateFilePath(addParams.ImagePath, "imagePath", true);

        if (!File.Exists(addParams.ImagePath))
            throw new FileNotFoundException($"Image file not found: {addParams.ImagePath}");

        ExcelImageHelper.ValidateImageFormat(addParams.ImagePath);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, addParams.SheetIndex);
        var cellObj = worksheet.Cells[addParams.Cell];

        var pictureIndex = worksheet.Pictures.Add(cellObj.Row, cellObj.Column, addParams.ImagePath);
        var picture = worksheet.Pictures[pictureIndex];

        if (addParams.Width.HasValue || addParams.Height.HasValue)
        {
            picture.IsLockAspectRatio = addParams.KeepAspectRatio;
            if (addParams.Width.HasValue) picture.Width = addParams.Width.Value;
            if (addParams.Height.HasValue) picture.Height = addParams.Height.Value;
        }

        MarkModified(context);

        return new SuccessResult
            { Message = $"Image added to cell {addParams.Cell} (size: {picture.Width}x{picture.Height})." };
    }

    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var imagePath = parameters.GetOptional<string?>("imagePath");
        var cell = parameters.GetOptional<string?>("cell");
        var width = parameters.GetOptional<int?>("width");
        var height = parameters.GetOptional<int?>("height");
        var keepAspectRatio = parameters.GetOptional("keepAspectRatio", true);

        if (string.IsNullOrEmpty(imagePath))
            throw new ArgumentException("imagePath is required for add operation");
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for add operation");

        return new AddParameters(sheetIndex, imagePath, cell, width, height, keepAspectRatio);
    }

    private sealed record AddParameters(
        int SheetIndex,
        string ImagePath,
        string Cell,
        int? Width,
        int? Height,
        bool KeepAspectRatio);
}
