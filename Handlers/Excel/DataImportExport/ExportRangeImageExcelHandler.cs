using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.DataImportExport;

namespace AsposeMcpServer.Handlers.Excel.DataImportExport;

/// <summary>
///     Handler for exporting an Excel range to an image.
/// </summary>
[ResultType(typeof(ExportExcelResult))]
public class ExportRangeImageExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "export_range_image";

    /// <summary>
    ///     Exports a worksheet or range to an image.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: outputPath
    ///     Optional: sheetIndex (default: 0), format (default: png), dpi (default: 150)
    /// </param>
    /// <returns>Export result with output path.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var format = parameters.GetOptional("format", "png");
        var dpi = parameters.GetOptional("dpi", 150);

        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for export_range_image operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var imageType = ResolveImageType(format);
            var options = new ImageOrPrintOptions
            {
                ImageType = imageType,
                HorizontalResolution = dpi,
                VerticalResolution = dpi,
                OnePagePerSheet = true
            };

            var sheetRender = new SheetRender(worksheet, options);

            if (sheetRender.PageCount > 0)
                sheetRender.ToImage(0, outputPath);

            return new ExportExcelResult
            {
                OutputPath = outputPath,
                RowCount = worksheet.Cells.MaxDataRow + 1,
                Message = $"Sheet {sheetIndex} exported to image: {outputPath} (format: {format}, DPI: {dpi})."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to export range image: {ex.Message}");
        }
    }

    /// <summary>
    ///     Resolves an image format string to an ImageType enum value.
    /// </summary>
    /// <param name="format">The image format string.</param>
    /// <returns>The corresponding ImageType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is unknown.</exception>
    internal static ImageType ResolveImageType(string format)
    {
        return format.ToLowerInvariant() switch
        {
            "png" => ImageType.Png,
            "jpeg" or "jpg" => ImageType.Jpeg,
            "bmp" => ImageType.Bmp,
            "tiff" or "tif" => ImageType.Tiff,
            "svg" => ImageType.Svg,
            "gif" => ImageType.Gif,
            "emf" => ImageType.Emf,
            _ => throw new ArgumentException(
                $"Unknown image format: '{format}'. Supported: png, jpeg, bmp, tiff, svg, gif, emf")
        };
    }
}
