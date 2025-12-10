using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetRangeTool : IAsposeTool
{
    public string Description => "Get range data with optional formatting information from Excel";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            range = new
            {
                type = "string",
                description = "Cell range (e.g., 'A1:C5')"
            },
            includeFormulas = new
            {
                type = "boolean",
                description = "Include formulas instead of values (optional, default: false)"
            },
            includeFormat = new
            {
                type = "boolean",
                description = "Include format information (optional, default: false)"
            }
        },
        required = new[] { "path", "range" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var includeFormulas = arguments?["includeFormulas"]?.GetValue<bool?>() ?? false;
        var includeFormat = arguments?["includeFormat"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        // Calculate formulas before reading to ensure values are up-to-date
        workbook.CalculateFormula();

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        var sb = new StringBuilder();
        sb.AppendLine($"Range: {range}");
        sb.AppendLine($"Rows: {cellRange.RowCount}, Columns: {cellRange.ColumnCount}");
        sb.AppendLine();

        for (int i = 0; i < cellRange.RowCount; i++)
        {
            for (int j = 0; j < cellRange.ColumnCount; j++)
            {
                var cell = cells[cellRange.FirstRow + i, cellRange.FirstColumn + j];
                var cellRef = CellsHelper.CellIndexToName(cellRange.FirstRow + i, cellRange.FirstColumn + j);
                
                if (includeFormulas && !string.IsNullOrEmpty(cell.Formula))
                {
                    sb.Append($"{cellRef}: {cell.Formula}");
                }
                else
                {
                    // If cell has a formula, get the calculated value
                    // Otherwise, get the cell value directly
                    object? displayValue;
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // For formula cells, get the calculated value
                        displayValue = cell.Value;
                        
                        // Check if value is an error (like #DIV/0!, #VALUE!, etc.)
                        if (displayValue is Aspose.Cells.CellValueType cellType && cellType == Aspose.Cells.CellValueType.IsError)
                        {
                            // Try to get display string value instead
                            displayValue = cell.DisplayStringValue;
                        }
                        
                        // If value is null or empty, try display string value
                        if (displayValue == null || (displayValue is string str && string.IsNullOrEmpty(str)))
                        {
                            displayValue = cell.DisplayStringValue;
                            // If still empty, show formula as fallback
                            if (string.IsNullOrEmpty(displayValue?.ToString()))
                            {
                                displayValue = cell.Formula;
                            }
                        }
                    }
                    else
                    {
                        displayValue = cell.Value ?? cell.DisplayStringValue;
                    }
                    
                    sb.Append($"{cellRef}: {displayValue ?? "(empty)"}");
                }

                if (includeFormat)
                {
                    var style = cell.GetStyle();
                    sb.Append($" [Font: {style.Font.Name}, Size: {style.Font.Size}]");
                }

                if (j < cellRange.ColumnCount - 1)
                {
                    sb.Append(" | ");
                }
            }
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

