using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Handler for deleting images from Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, deleteParams.SheetIndex);
        var pictures = worksheet.Pictures;

        ExcelImageHelper.ValidateImageIndex(deleteParams.ImageIndex, pictures.Count);

        pictures.RemoveAt(deleteParams.ImageIndex);

        MarkModified(context);

        var warning = pictures.Count > 0
            ? " Note: remaining image indices have been re-ordered."
            : "";
        return new SuccessResult
            { Message = $"Image #{deleteParams.ImageIndex} deleted. {pictures.Count} images remaining.{warning}" };
    }

    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var imageIndex = parameters.GetOptional<int?>("imageIndex");

        if (!imageIndex.HasValue)
            throw new ArgumentException("imageIndex is required for delete operation");

        return new DeleteParameters(sheetIndex, imageIndex.Value);
    }

    private sealed record DeleteParameters(int SheetIndex, int ImageIndex);
}
