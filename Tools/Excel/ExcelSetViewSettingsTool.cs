using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetViewSettingsTool : IAsposeTool
{
    public string Description => "Set worksheet view settings (zoom, gridlines, headings, etc.)";

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
            zoom = new
            {
                type = "number",
                description = "Zoom percentage (optional, 10-400)"
            },
            showGridlines = new
            {
                type = "boolean",
                description = "Show gridlines (optional)"
            },
            showRowColumnHeaders = new
            {
                type = "boolean",
                description = "Show row/column headers (optional)"
            },
            showZeroValues = new
            {
                type = "boolean",
                description = "Show zero values (optional)"
            },
            displayRightToLeft = new
            {
                type = "boolean",
                description = "Display right to left (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var zoom = arguments?["zoom"]?.GetValue<int?>();
        var showGridlines = arguments?["showGridlines"]?.GetValue<bool?>();
        var showRowColumnHeaders = arguments?["showRowColumnHeaders"]?.GetValue<bool?>();
        var showZeroValues = arguments?["showZeroValues"]?.GetValue<bool?>();
        var displayRightToLeft = arguments?["displayRightToLeft"]?.GetValue<bool?>();

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];

        if (zoom.HasValue)
        {
            if (zoom.Value < 10 || zoom.Value > 400)
            {
                throw new ArgumentException("Zoom must be between 10 and 400");
            }
            worksheet.Zoom = zoom.Value;
        }

        if (showGridlines.HasValue)
        {
            worksheet.IsGridlinesVisible = showGridlines.Value;
        }

        if (showRowColumnHeaders.HasValue)
        {
            worksheet.IsRowColumnHeadersVisible = showRowColumnHeaders.Value;
        }

        if (showZeroValues.HasValue)
        {
            worksheet.DisplayZeros = showZeroValues.Value;
        }

        if (displayRightToLeft.HasValue)
        {
            worksheet.DisplayRightToLeft = displayRightToLeft.Value;
        }

        workbook.Save(path);
        return await Task.FromResult($"View settings updated for sheet {sheetIndex}: {path}");
    }
}

