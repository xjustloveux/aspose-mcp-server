using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Handler for batch writing multiple values to Excel cells.
/// </summary>
public class BatchWriteHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "batch_write";

    /// <summary>
    ///     Writes multiple values to cells in batch.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex, data (JSON array or object)
    /// </param>
    /// <returns>Success message with write count.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var data = parameters.GetOptional<JsonNode?>("data");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var writeCount = 0;

            if (data != null)
            {
                if (data is JsonArray dataArray)
                    foreach (var item in dataArray)
                    {
                        var itemObj = item?.AsObject();
                        if (itemObj != null)
                        {
                            var cell = itemObj["cell"]?.GetValue<string>();
                            var value = itemObj["value"]?.GetValue<string>() ?? "";
                            if (!string.IsNullOrEmpty(cell))
                            {
                                var cellObj = worksheet.Cells[cell];
                                ExcelHelper.SetCellValue(cellObj, value);
                                writeCount++;
                            }
                        }
                    }
                else if (data is JsonObject dataObject)
                    foreach (var kvp in dataObject)
                    {
                        var cell = kvp.Key;
                        var value = kvp.Value?.GetValue<string>() ?? "";
                        if (!string.IsNullOrEmpty(cell))
                        {
                            var cellObj = worksheet.Cells[cell];
                            ExcelHelper.SetCellValue(cellObj, value);
                            writeCount++;
                        }
                    }
            }

            MarkModified(context);

            return Success($"Batch write completed ({writeCount} cells written to sheet {sheetIndex}).");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }
}
