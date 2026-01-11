using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Handler for deleting images from Excel worksheets.
/// </summary>
public class DeleteExcelImageHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes an image from the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: imageIndex
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var imageIndex = parameters.GetRequired<int>("imageIndex");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pictures = worksheet.Pictures;

        ExcelImageHelper.ValidateImageIndex(imageIndex, pictures.Count);

        pictures.RemoveAt(imageIndex);

        MarkModified(context);

        var warning = pictures.Count > 0
            ? " Note: remaining image indices have been re-ordered."
            : "";
        return Success($"Image #{imageIndex} deleted. {pictures.Count} images remaining.{warning}");
    }
}
