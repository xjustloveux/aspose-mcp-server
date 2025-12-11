using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel print settings (print area, titles, page setup, etc.)
/// Merges: ExcelSetPrintAreaTool, ExcelSetPrintTitlesTool, ExcelSetPrintSettingsTool, ExcelSetPageSetupTool
/// </summary>
public class ExcelPrintSettingsTool : IAsposeTool
{
    public string Description => "Manage Excel print settings: set print area, titles, page setup, or print settings";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'set_print_area', 'set_print_titles', 'set_page_setup', 'set_all'",
                @enum = new[] { "set_print_area", "set_print_titles", "set_page_setup", "set_all" }
            },
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
                description = "Print area range (e.g., 'A1:D10', optional for set_print_area)"
            },
            clearPrintArea = new
            {
                type = "boolean",
                description = "Clear print area (optional, for set_print_area, default: false)"
            },
            rows = new
            {
                type = "string",
                description = "Rows to repeat (e.g., '1:1', optional for set_print_titles)"
            },
            columns = new
            {
                type = "string",
                description = "Columns to repeat (e.g., 'A:A', optional for set_print_titles)"
            },
            clearTitles = new
            {
                type = "boolean",
                description = "Clear print titles (optional, for set_print_titles, default: false)"
            },
            orientation = new
            {
                type = "string",
                description = "Page orientation: 'Portrait' or 'Landscape' (optional)",
                @enum = new[] { "Portrait", "Landscape" }
            },
            paperSize = new
            {
                type = "string",
                description = "Paper size (e.g., 'A4', 'Letter', 'Legal', optional)"
            },
            leftMargin = new
            {
                type = "number",
                description = "Left margin in inches (optional)"
            },
            rightMargin = new
            {
                type = "number",
                description = "Right margin in inches (optional)"
            },
            topMargin = new
            {
                type = "number",
                description = "Top margin in inches (optional)"
            },
            bottomMargin = new
            {
                type = "number",
                description = "Bottom margin in inches (optional)"
            },
            header = new
            {
                type = "string",
                description = "Header text (optional)"
            },
            footer = new
            {
                type = "string",
                description = "Footer text (optional)"
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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "set_print_area" => await SetPrintAreaAsync(arguments, path, sheetIndex),
            "set_print_titles" => await SetPrintTitlesAsync(arguments, path, sheetIndex),
            "set_page_setup" => await SetPageSetupAsync(arguments, path, sheetIndex),
            "set_all" => await SetAllAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> SetPrintAreaAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>();
        var clearPrintArea = arguments?["clearPrintArea"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (clearPrintArea)
        {
            worksheet.PageSetup.PrintArea = "";
        }
        else if (!string.IsNullOrEmpty(range))
        {
            worksheet.PageSetup.PrintArea = range;
        }
        else
        {
            throw new ArgumentException("Either range or clearPrintArea must be provided");
        }

        workbook.Save(path);
        return await Task.FromResult(clearPrintArea
            ? $"Print area cleared for sheet {sheetIndex}: {path}"
            : $"Print area set to {range} for sheet {sheetIndex}: {path}");
    }

    private async Task<string> SetPrintTitlesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var rows = arguments?["rows"]?.GetValue<string>();
        var columns = arguments?["columns"]?.GetValue<string>();
        var clearTitles = arguments?["clearTitles"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (clearTitles)
        {
            worksheet.PageSetup.PrintTitleRows = "";
            worksheet.PageSetup.PrintTitleColumns = "";
        }
        else
        {
            if (!string.IsNullOrEmpty(rows))
            {
                worksheet.PageSetup.PrintTitleRows = rows;
            }
            if (!string.IsNullOrEmpty(columns))
            {
                worksheet.PageSetup.PrintTitleColumns = columns;
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Print titles updated for sheet {sheetIndex}: {path}");
    }

    private async Task<string> SetPageSetupAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var orientation = arguments?["orientation"]?.GetValue<string>();
        var paperSize = arguments?["paperSize"]?.GetValue<string>();
        var leftMargin = arguments?["leftMargin"]?.GetValue<double?>();
        var rightMargin = arguments?["rightMargin"]?.GetValue<double?>();
        var topMargin = arguments?["topMargin"]?.GetValue<double?>();
        var bottomMargin = arguments?["bottomMargin"]?.GetValue<double?>();
        var header = arguments?["header"]?.GetValue<string>();
        var footer = arguments?["footer"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var pageSetup = worksheet.PageSetup;

        if (!string.IsNullOrEmpty(orientation))
        {
            pageSetup.Orientation = orientation == "Landscape" ? PageOrientationType.Landscape : PageOrientationType.Portrait;
        }

        if (!string.IsNullOrEmpty(paperSize))
        {
            var size = paperSize.ToUpper() switch
            {
                "A4" => PaperSizeType.PaperA4,
                "LETTER" => PaperSizeType.PaperLetter,
                "LEGAL" => PaperSizeType.PaperLegal,
                "A3" => PaperSizeType.PaperA3,
                "A5" => PaperSizeType.PaperA5,
                "B4" => PaperSizeType.PaperB4,
                "B5" => PaperSizeType.PaperB5,
                _ => PaperSizeType.PaperA4
            };
            pageSetup.PaperSize = size;
        }

        if (leftMargin.HasValue) pageSetup.LeftMargin = leftMargin.Value;
        if (rightMargin.HasValue) pageSetup.RightMargin = rightMargin.Value;
        if (topMargin.HasValue) pageSetup.TopMargin = topMargin.Value;
        if (bottomMargin.HasValue) pageSetup.BottomMargin = bottomMargin.Value;

        if (!string.IsNullOrEmpty(header))
        {
            pageSetup.SetHeader(0, header);
        }

        if (!string.IsNullOrEmpty(footer))
        {
            pageSetup.SetFooter(0, footer);
        }

        workbook.Save(path);

        var changes = new List<string>();
        if (!string.IsNullOrEmpty(orientation)) changes.Add($"方向: {orientation}");
        if (!string.IsNullOrEmpty(paperSize)) changes.Add($"紙張大小: {paperSize}");
        if (leftMargin.HasValue || rightMargin.HasValue || topMargin.HasValue || bottomMargin.HasValue)
            changes.Add("邊距已設定");
        if (!string.IsNullOrEmpty(header)) changes.Add("頁首已設定");
        if (!string.IsNullOrEmpty(footer)) changes.Add("頁尾已設定");

        return await Task.FromResult($"頁面設定已更新: {string.Join(", ", changes)}");
    }

    private async Task<string> SetAllAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var printArea = arguments?["range"]?.GetValue<string>();
        var printTitleRows = arguments?["rows"]?.GetValue<string>();
        var printTitleColumns = arguments?["columns"]?.GetValue<string>();
        var fitToPage = arguments?["fitToPage"]?.GetValue<bool?>();
        var fitToPagesWide = arguments?["fitToPagesWide"]?.GetValue<int?>();
        var fitToPagesTall = arguments?["fitToPagesTall"]?.GetValue<int?>();
        var orientation = arguments?["orientation"]?.GetValue<string>();
        var paperSize = arguments?["paperSize"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
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

