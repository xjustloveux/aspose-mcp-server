using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Render;

namespace AsposeMcpServer.Handlers.Excel.Render;

/// <summary>
///     Handler for rendering an Excel worksheet to an image.
/// </summary>
[ResultType(typeof(RenderExcelResult))]
public class RenderSheetExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "render_sheet";

    /// <summary>
    ///     Renders a worksheet to image(s).
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: outputPath
    ///     Optional: sheetIndex (default: 0), format (default: png), dpi (default: 150)
    /// </param>
    /// <returns>Render result with output paths.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var format = parameters.GetOptional("format", "png");
        var dpi = parameters.GetOptional("dpi", 150);

        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for render_sheet operation");

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
            var outputPaths = new List<string>();

            if (sheetRender.PageCount <= 1)
            {
                if (sheetRender.PageCount > 0)
                    sheetRender.ToImage(0, outputPath);
                outputPaths.Add(outputPath);
            }
            else
            {
                var dir = Path.GetDirectoryName(outputPath) ?? ".";
                var nameWithoutExt = Path.GetFileNameWithoutExtension(outputPath);
                var ext = Path.GetExtension(outputPath);

                for (var i = 0; i < sheetRender.PageCount; i++)
                {
                    var pagePath = Path.Combine(dir, $"{nameWithoutExt}_{i + 1}{ext}");
                    sheetRender.ToImage(i, pagePath);
                    outputPaths.Add(pagePath);
                }
            }

            return new RenderExcelResult
            {
                OutputPaths = outputPaths,
                PageCount = sheetRender.PageCount,
                Format = format,
                Message =
                    $"Sheet {sheetIndex} rendered to {sheetRender.PageCount} page(s) in {format} format at {dpi} DPI."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to render sheet: {ex.Message}");
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
