using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Render;

namespace AsposeMcpServer.Handlers.Excel.Render;

/// <summary>
///     Handler for rendering an Excel chart to an image.
/// </summary>
[ResultType(typeof(RenderExcelResult))]
public class RenderChartExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "render_chart";

    /// <summary>
    ///     Renders a chart to an image.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: outputPath, chartIndex
    ///     Optional: sheetIndex (default: 0), format (default: png)
    /// </param>
    /// <returns>Render result with output path.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var chartIndex = parameters.GetOptional<int?>("chartIndex");
        var format = parameters.GetOptional("format", "png");

        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for render_chart operation");
        if (!chartIndex.HasValue)
            throw new ArgumentException("chartIndex is required for render_chart operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (chartIndex.Value < 0 || chartIndex.Value >= worksheet.Charts.Count)
                throw new ArgumentException(
                    $"Chart index {chartIndex.Value} is out of range (worksheet has {worksheet.Charts.Count} charts)");

            var chart = worksheet.Charts[chartIndex.Value];
            var imageType = RenderSheetExcelHandler.ResolveImageType(format);

            chart.ToImage(outputPath, imageType);

            return new RenderExcelResult
            {
                OutputPaths = new List<string> { outputPath },
                PageCount = 1,
                Format = format,
                Message =
                    $"Chart {chartIndex.Value} from sheet {sheetIndex} rendered to {outputPath} in {format} format."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to render chart: {ex.Message}");
        }
    }
}
