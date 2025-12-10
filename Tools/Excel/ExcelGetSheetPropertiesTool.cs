using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetSheetPropertiesTool : IAsposeTool
{
    public string Description => "Get worksheet properties and settings from Excel";

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
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var sb = new StringBuilder();

        sb.AppendLine($"Sheet Properties:");
        sb.AppendLine($"  Name: {worksheet.Name}");
        sb.AppendLine($"  Index: {sheetIndex}");
        sb.AppendLine($"  Is Visible: {worksheet.IsVisible}");
        sb.AppendLine($"  Tab Color: {worksheet.TabColor}");
        sb.AppendLine($"  Is Selected: {workbook.Worksheets.ActiveSheetIndex == sheetIndex}");
        sb.AppendLine($"  Max Data Row: {worksheet.Cells.MaxDataRow}");
        sb.AppendLine($"  Max Data Column: {worksheet.Cells.MaxDataColumn}");
        sb.AppendLine($"  Is Protected: {worksheet.Protection.IsProtectedWithPassword}");
        sb.AppendLine($"  Comments Count: {worksheet.Comments.Count}");
        sb.AppendLine($"  Charts Count: {worksheet.Charts.Count}");
        sb.AppendLine($"  Pictures Count: {worksheet.Pictures.Count}");
        sb.AppendLine($"  Hyperlinks Count: {worksheet.Hyperlinks.Count}");

        var pageSetup = worksheet.PageSetup;
        sb.AppendLine($"\nPrint Settings:");
        sb.AppendLine($"  Print Area: {pageSetup.PrintArea ?? "(none)"}");
        sb.AppendLine($"  Print Title Rows: {pageSetup.PrintTitleRows ?? "(none)"}");
        sb.AppendLine($"  Print Title Columns: {pageSetup.PrintTitleColumns ?? "(none)"}");
        sb.AppendLine($"  Orientation: {pageSetup.Orientation}");
        sb.AppendLine($"  Paper Size: {pageSetup.PaperSize}");
        sb.AppendLine($"  Fit To Pages Wide: {pageSetup.FitToPagesWide}");
        sb.AppendLine($"  Fit To Pages Tall: {pageSetup.FitToPagesTall}");

        return await Task.FromResult(sb.ToString());
    }
}

