using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Handler for batch writing multiple values to Excel cells.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var batchWriteParams = ExtractBatchWriteParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, batchWriteParams.SheetIndex);

            var writeCount = WriteData(worksheet, batchWriteParams.Data);

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Batch write completed ({writeCount} cells written to sheet {batchWriteParams.SheetIndex})."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    /// <summary>
    ///     Extracts batch write parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted batch write parameters.</returns>
    private static BatchWriteParameters ExtractBatchWriteParameters(OperationParameters parameters)
    {
        return new BatchWriteParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<JsonNode?>("data")
        );
    }

    /// <summary>
    ///     Writes data to a worksheet from a JSON node.
    /// </summary>
    /// <param name="worksheet">The worksheet to write to.</param>
    /// <param name="data">The JSON node containing the data to write.</param>
    /// <returns>The number of cells written.</returns>
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

    /// <summary>
    ///     Writes data from a JSON array to a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to write to.</param>
    /// <param name="dataArray">The JSON array containing cell-value pairs.</param>
    /// <returns>The number of cells written.</returns>
    private static int WriteFromArray(Worksheet worksheet, JsonArray dataArray)
    {
        return dataArray.Count(item => WriteArrayItem(worksheet, item));
    }

    /// <summary>
    ///     Writes a single array item to the worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to write to.</param>
    /// <param name="item">The JSON node containing cell and value properties.</param>
    /// <returns>True if the item was written successfully; otherwise, false.</returns>
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

    /// <summary>
    ///     Writes data from a JSON object to a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to write to.</param>
    /// <param name="dataObject">The JSON object with cell addresses as keys and values as values.</param>
    /// <returns>The number of cells written.</returns>
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

    /// <summary>
    ///     Parameters for batch write operation.
    /// </summary>
    /// <param name="SheetIndex">The worksheet index (0-based).</param>
    /// <param name="Data">The JSON data containing cell-value pairs to write.</param>
    private sealed record BatchWriteParameters(int SheetIndex, JsonNode? Data);
}
