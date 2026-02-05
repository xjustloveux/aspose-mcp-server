using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.DataImportExport;

namespace AsposeMcpServer.Handlers.Excel.DataImportExport;

/// <summary>
///     Handler for exporting Excel worksheet data to CSV format.
/// </summary>
[ResultType(typeof(ExportExcelResult))]
public class ExportCsvExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "export_csv";

    /// <summary>
    ///     Exports worksheet data to a CSV file.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: outputPath
    ///     Optional: sheetIndex (default: 0), separator (default: ',')
    /// </param>
    /// <returns>Export result with output path.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var separator = parameters.GetOptional("separator", ",");

        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for export_csv operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        try
        {
            var workbook = context.Document;
            ExcelHelper.ValidateSheetIndex(sheetIndex, workbook);

            var saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = separator.Length > 0 ? separator[0] : ','
            };

            workbook.Worksheets.ActiveSheetIndex = sheetIndex;
            workbook.Save(outputPath, saveOptions);

            var worksheet = workbook.Worksheets[sheetIndex];
            var rowCount = worksheet.Cells.MaxDataRow + 1;

            return new ExportExcelResult
            {
                OutputPath = outputPath,
                RowCount = rowCount,
                Message = $"Sheet {sheetIndex} exported to CSV: {outputPath} ({rowCount} rows)."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to export to CSV: {ex.Message}");
        }
    }
}
