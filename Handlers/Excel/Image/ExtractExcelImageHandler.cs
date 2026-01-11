using Aspose.Cells;
using Aspose.Cells.Rendering;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Handler for extracting images from Excel worksheets.
/// </summary>
public class ExtractExcelImageHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>
    ///     Extracts an image from the worksheet and saves it to a file.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: imageIndex, exportPath
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with extraction details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var imageIndex = parameters.GetRequired<int>("imageIndex");
        var exportPath = parameters.GetRequired<string>("exportPath");

        SecurityHelper.ValidateFilePath(exportPath, "exportPath", true);

        var extension = Path.GetExtension(exportPath);
        if (string.IsNullOrEmpty(extension) ||
            !ExcelImageHelper.ExtensionToImageType.TryGetValue(extension, out var imageType))
            throw new ArgumentException(
                $"Unsupported export format: '{extension}'. Supported formats: {string.Join(", ", ExcelImageHelper.ExtensionToImageType.Keys)}");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pictures = worksheet.Pictures;

        ExcelImageHelper.ValidateImageIndex(imageIndex, pictures.Count);

        var picture = pictures[imageIndex];
        var upperLeftCell = CellsHelper.CellIndexToName(picture.UpperLeftRow, picture.UpperLeftColumn);

        var exportDir = Path.GetDirectoryName(exportPath);
        if (!string.IsNullOrEmpty(exportDir) && !Directory.Exists(exportDir))
            Directory.CreateDirectory(exportDir);

        var options = new ImageOrPrintOptions
        {
            ImageType = imageType
        };
        picture.ToImage(exportPath, options);

        var fileInfo = new FileInfo(exportPath);
        return Success(
            $"Image #{imageIndex} (at {upperLeftCell}) extracted to: {exportPath} ({fileInfo.Length} bytes, {picture.Width}x{picture.Height})");
    }
}
