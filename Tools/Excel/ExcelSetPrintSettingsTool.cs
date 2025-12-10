using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetPrintSettingsTool : IAsposeTool
{
    public string Description => "Set print settings for Excel worksheet";

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
            printArea = new
            {
                type = "string",
                description = "Print area range (e.g., 'A1:D10', optional)"
            },
            printTitleRows = new
            {
                type = "string",
                description = "Rows to repeat at top (e.g., '1:1', optional)"
            },
            printTitleColumns = new
            {
                type = "string",
                description = "Columns to repeat at left (e.g., 'A:A', optional)"
            },
            fitToPage = new
            {
                type = "boolean",
                description = "Fit to page (optional)"
            },
            fitToPagesWide = new
            {
                type = "number",
                description = "Fit to pages wide (optional)"
            },
            fitToPagesTall = new
            {
                type = "number",
                description = "Fit to pages tall (optional)"
            },
            orientation = new
            {
                type = "string",
                description = "Orientation (Portrait or Landscape, optional)"
            },
            paperSize = new
            {
                type = "string",
                description = "Paper size (A4, Letter, etc., optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var printArea = arguments?["printArea"]?.GetValue<string>();
        var printTitleRows = arguments?["printTitleRows"]?.GetValue<string>();
        var printTitleColumns = arguments?["printTitleColumns"]?.GetValue<string>();
        var fitToPage = arguments?["fitToPage"]?.GetValue<bool?>();
        var fitToPagesWide = arguments?["fitToPagesWide"]?.GetValue<int?>();
        var fitToPagesTall = arguments?["fitToPagesTall"]?.GetValue<int?>();
        var orientation = arguments?["orientation"]?.GetValue<string>();
        var paperSize = arguments?["paperSize"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var pageSetup = worksheet.PageSetup;

        if (!string.IsNullOrEmpty(printArea))
        {
            pageSetup.PrintArea = printArea;
        }

        if (!string.IsNullOrEmpty(printTitleRows))
        {
            pageSetup.PrintTitleRows = printTitleRows;
        }

        if (!string.IsNullOrEmpty(printTitleColumns))
        {
            pageSetup.PrintTitleColumns = printTitleColumns;
        }

        if (fitToPage.HasValue)
        {
            pageSetup.FitToPagesWide = fitToPagesWide ?? 1;
            pageSetup.FitToPagesTall = fitToPagesTall ?? 1;
        }

        if (!string.IsNullOrEmpty(orientation))
        {
            pageSetup.Orientation = orientation.ToLower() == "landscape" ? PageOrientationType.Landscape : PageOrientationType.Portrait;
        }

        if (!string.IsNullOrEmpty(paperSize))
        {
            var paperSizeEnum = paperSize.ToUpper() switch
            {
                "A4" => PaperSizeType.PaperA4,
                "LETTER" => PaperSizeType.PaperLetter,
                "LEGAL" => PaperSizeType.PaperLegal,
                "A3" => PaperSizeType.PaperA3,
                "A5" => PaperSizeType.PaperA5,
                _ => PaperSizeType.PaperA4
            };
            pageSetup.PaperSize = paperSizeEnum;
        }

        workbook.Save(path);
        return await Task.FromResult($"Print settings updated for sheet {sheetIndex}: {path}");
    }
}

