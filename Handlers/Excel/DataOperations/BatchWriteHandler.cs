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

            var writeCount = WriteData(worksheet, data);

            MarkModified(context);

            return Success($"Batch write completed ({writeCount} cells written to sheet {sheetIndex}).");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    private static int WriteData(Worksheet worksheet, JsonNode? data)
    {
        if (data == null) return 0;

        return data switch
        {
            JsonArray dataArray => WriteFromArray(worksheet, dataArray),
            JsonObject dataObject => WriteFromObject(worksheet, dataObject),
            _ => 0
        };
    }

    private static int WriteFromArray(Worksheet worksheet, JsonArray dataArray)
    {
        var writeCount = 0;
        foreach (var item in dataArray)
            if (WriteArrayItem(worksheet, item))
                writeCount++;

        return writeCount;
    }

    private static bool WriteArrayItem(Worksheet worksheet, JsonNode? item)
    {
        var itemObj = item?.AsObject();
        if (itemObj == null) return false;

        var cell = itemObj["cell"]?.GetValue<string>();
        if (string.IsNullOrEmpty(cell)) return false;

        var value = itemObj["value"]?.GetValue<string>() ?? "";
        var cellObj = worksheet.Cells[cell];
        ExcelHelper.SetCellValue(cellObj, value);
        return true;
    }

    private static int WriteFromObject(Worksheet worksheet, JsonObject dataObject)
    {
        var writeCount = 0;
        foreach (var kvp in dataObject)
        {
            var cell = kvp.Key;
            if (string.IsNullOrEmpty(cell)) continue;

            var value = kvp.Value?.GetValue<string>() ?? "";
            var cellObj = worksheet.Cells[cell];
            ExcelHelper.SetCellValue(cellObj, value);
            writeCount++;
        }

        return writeCount;
    }
}
