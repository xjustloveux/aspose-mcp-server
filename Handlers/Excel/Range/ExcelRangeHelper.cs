using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Helper class for Excel range operations.
/// </summary>
public static class ExcelRangeHelper
{
    /// <summary>
    ///     Text format number for Aspose.Cells style (formats cell as text).
    /// </summary>
    public const int TextFormatNumber = 49;

    /// <summary>
    ///     Parses and validates JSON data array.
    /// </summary>
    /// <param name="dataJson">The JSON string containing the data.</param>
    /// <returns>The parsed JsonArray.</returns>
    /// <exception cref="ArgumentException">Thrown when the JSON is invalid or not an array.</exception>
    public static JsonArray ParseDataArray(string dataJson)
    {
        try
        {
            var parsed = JsonNode.Parse(dataJson);
            return parsed?.AsArray() ?? throw new ArgumentException("data must be a JSON array");
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid JSON format for data: {ex.Message}");
        }
    }

    /// <summary>
    ///     Checks if a cell value looks like a cell reference.
    /// </summary>
    /// <param name="cellValue">The value to check.</param>
    /// <returns>True if the value looks like a cell reference.</returns>
    public static bool LooksLikeCellReference(string cellValue)
    {
        return cellValue.Length >= 2 &&
               char.IsLetter(cellValue[0]) &&
               ((cellValue.Length is 2 && char.IsDigit(cellValue[1])) ||
                (cellValue.Length is > 2 and <= 5 &&
                 cellValue.Skip(1).All(char.IsLetterOrDigit) &&
                 cellValue.Substring(1).Any(char.IsDigit) &&
                 !cellValue.Contains(' ') &&
                 !cellValue.Contains(':') &&
                 !cellValue.Contains('$')));
    }

    /// <summary>
    ///     Writes 2D array data to worksheet starting at specified cell.
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <param name="worksheet">The worksheet to write to.</param>
    /// <param name="startRow">The starting row index.</param>
    /// <param name="startCol">The starting column index.</param>
    /// <param name="dataArray">The 2D data array.</param>
    public static void Write2DArrayData(Workbook workbook, Worksheet worksheet, int startRow, int startCol,
        JsonArray dataArray)
    {
        var rowCount = dataArray.Count;
        var colCount = dataArray.Max(item => item is JsonArray arr ? arr.Count : 0);

        if (colCount == 0) return;

        var data2D = new object[rowCount, colCount];

        for (var i = 0; i < rowCount; i++)
            if (dataArray[i] is JsonArray rowArray)
                for (var j = 0; j < colCount; j++)
                    if (j < rowArray.Count)
                        data2D[i, j] = ValueHelper.ParseValue(rowArray[j]?.GetValue<string>() ?? "");
                    else
                        data2D[i, j] = "";

        worksheet.Cells.ImportTwoDimensionArray(data2D, startRow, startCol);

        // Handle potential cell references
        for (var i = 0; i < rowCount; i++)
            if (dataArray[i] is JsonArray rowArray)
                for (var j = 0; j < rowArray.Count; j++)
                {
                    var cellValue = rowArray[j]?.GetValue<string>() ?? "";

                    if (LooksLikeCellReference(cellValue) && ValueHelper.ParseValue(cellValue) is string)
                    {
                        var cellObj = worksheet.Cells[startRow + i, startCol + j];
                        var style = workbook.CreateStyle();
                        style.Number = TextFormatNumber;
                        cellObj.SetStyle(style);
                        cellObj.PutValue(cellValue, true);
                    }
                }
    }

    /// <summary>
    ///     Writes object array data (with cell and value properties) to worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to write to.</param>
    /// <param name="dataArray">The object array data.</param>
    /// <exception cref="ArgumentException">Thrown when data format is invalid.</exception>
    public static void WriteObjectArrayData(Worksheet worksheet, JsonArray dataArray)
    {
        for (var i = 0; i < dataArray.Count; i++)
        {
            var item = dataArray[i];

            if (item is JsonObject itemObj)
            {
                var cellRef = itemObj["cell"]?.GetValue<string>();
                var cellValue = itemObj["value"]?.GetValue<string>() ?? "";

                if (!string.IsNullOrEmpty(cellRef))
                    ExcelHelper.SetCellValue(worksheet.Cells[cellRef], cellValue);
            }
            else
            {
                throw new ArgumentException(
                    $"Invalid data format at index {i}. Expected array of arrays (2D) or array of objects with 'cell' and 'value' properties. Got: {item?.GetType().Name ?? "null"}");
            }
        }
    }

    /// <summary>
    ///     Converts copy options string to PasteType enum.
    /// </summary>
    /// <param name="copyOptions">The copy options string.</param>
    /// <returns>The PasteType enum value.</returns>
    public static PasteType GetPasteType(string copyOptions)
    {
        return copyOptions switch
        {
            "Values" => PasteType.Values,
            "Formats" => PasteType.Formats,
            "Formulas" => PasteType.Formulas,
            _ => PasteType.All
        };
    }
}
