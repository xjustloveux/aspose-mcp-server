using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetDataValidationInputMessageTool : IAsposeTool
{
    public string Description => "Set data validation input message (tooltip) in Excel";

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
                description = "Range with data validation (e.g., 'A1:A10')"
            },
            title = new
            {
                type = "string",
                description = "Input message title (optional)"
            },
            message = new
            {
                type = "string",
                description = "Input message text (optional)"
            },
            showInput = new
            {
                type = "boolean",
                description = "Show input message (optional, default: true)"
            }
        },
        required = new[] { "path", "range" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var title = arguments?["title"]?.GetValue<string>();
        var message = arguments?["message"]?.GetValue<string>();
        var showInput = arguments?["showInput"]?.GetValue<bool?>() ?? true;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var rangeObj = worksheet.Cells.CreateRange(range);

        // Find validation for the range by checking all validations
        Validation? validation = null;
        var rangeArea = new CellArea
        {
            StartRow = rangeObj.FirstRow,
            StartColumn = rangeObj.FirstColumn,
            EndRow = rangeObj.FirstRow + rangeObj.RowCount - 1,
            EndColumn = rangeObj.FirstColumn + rangeObj.ColumnCount - 1
        };
        
        for (int i = 0; i < worksheet.Validations.Count; i++)
        {
            var val = worksheet.Validations[i];
            // Check if validation covers the range (simplified check)
            if (val.Formula1 != null || val.Formula2 != null)
            {
                validation = val;
                break; // Use first matching validation
            }
        }

        if (validation == null)
        {
            throw new ArgumentException($"No data validation found for range {range}");
        }

        if (!string.IsNullOrEmpty(title))
        {
            validation.InputTitle = title;
        }
        if (!string.IsNullOrEmpty(message))
        {
            validation.InputMessage = message;
        }

        validation.ShowInput = showInput;

        workbook.Save(path);
        return await Task.FromResult($"Data validation input message updated for range {range}: {path}");
    }
}

