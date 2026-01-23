using Aspose.Cells;
using Aspose.Cells.Rendering;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Handler for extracting images from Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var extractParams = ExtractExtractParameters(parameters);

        SecurityHelper.ValidateFilePath(extractParams.ExportPath, "exportPath", true);

        var extension = Path.GetExtension(extractParams.ExportPath);
        if (string.IsNullOrEmpty(extension) ||
            !ExcelImageHelper.ExtensionToImageType.TryGetValue(extension, out var imageType))
            throw new ArgumentException(
                $"Unsupported export format: '{extension}'. Supported formats: {string.Join(", ", ExcelImageHelper.ExtensionToImageType.Keys)}");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, extractParams.SheetIndex);
        var pictures = worksheet.Pictures;

        ExcelImageHelper.ValidateImageIndex(extractParams.ImageIndex, pictures.Count);

        var picture = pictures[extractParams.ImageIndex];
        var upperLeftCell = CellsHelper.CellIndexToName(picture.UpperLeftRow, picture.UpperLeftColumn);

        var exportDir = Path.GetDirectoryName(extractParams.ExportPath);
        if (!string.IsNullOrEmpty(exportDir) && !Directory.Exists(exportDir))
            Directory.CreateDirectory(exportDir);

        var options = new ImageOrPrintOptions
        {
            ImageType = imageType
        };
        picture.ToImage(extractParams.ExportPath, options);

        var fileInfo = new FileInfo(extractParams.ExportPath);
        return new SuccessResult
        {
            Message =
                $"Image #{extractParams.ImageIndex} (at {upperLeftCell}) extracted to: {extractParams.ExportPath} ({fileInfo.Length} bytes, {picture.Width}x{picture.Height})"
        };
    }

    private static ExtractParameters ExtractExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var imageIndex = parameters.GetOptional<int?>("imageIndex");
        var exportPath = parameters.GetOptional<string?>("exportPath");

        if (!imageIndex.HasValue)
            throw new ArgumentException("imageIndex is required for extract operation");
        if (string.IsNullOrEmpty(exportPath))
            throw new ArgumentException("exportPath is required for extract operation");

        return new ExtractParameters(sheetIndex, imageIndex.Value, exportPath);
    }

    private sealed record ExtractParameters(int SheetIndex, int ImageIndex, string ExportPath);
}
