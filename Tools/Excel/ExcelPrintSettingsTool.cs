using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel print settings (print area, titles, page setup, etc.)
///     Merges: ExcelSetPrintAreaTool, ExcelSetPrintTitlesTool, ExcelSetPrintSettingsTool, ExcelSetPageSetupTool
/// </summary>
public class ExcelPrintSettingsTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
        @"Manage Excel print settings. Supports 4 operations: set_print_area, set_print_titles, set_page_setup, set_all.

Usage examples:
- Set print area: excel_print_settings(operation='set_print_area', path='book.xlsx', range='A1:D10')
- Set print titles: excel_print_settings(operation='set_print_titles', path='book.xlsx', rows='1:1', columns='A:A')
- Set page setup: excel_print_settings(operation='set_page_setup', path='book.xlsx', orientation='landscape', paperSize='A4')
- Set all: excel_print_settings(operation='set_all', path='book.xlsx', range='A1:D10', orientation='portrait')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set_print_area': Set print area (required params: path, range)
- 'set_print_titles': Set print titles (required params: path)
- 'set_page_setup': Set page setup (required params: path)
- 'set_all': Set all print settings (required params: path)",
                @enum = new[] { "set_print_area", "set_print_titles", "set_page_setup", "set_all" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
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
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for all operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "set_print_area" => await SetPrintAreaAsync(arguments, path, sheetIndex),
            "set_print_titles" => await SetPrintTitlesAsync(arguments, path, sheetIndex),
            "set_page_setup" => await SetPageSetupAsync(arguments, path, sheetIndex),
            "set_all" => await SetAllAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets print area for the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional range, clearPrintArea</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> SetPrintAreaAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetStringNullable(arguments, "range");
        var clearPrintArea = ArgumentHelper.GetBool(arguments, "clearPrintArea", false);

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (clearPrintArea)
            worksheet.PageSetup.PrintArea = "";
        else if (!string.IsNullOrEmpty(range))
            worksheet.PageSetup.PrintArea = range;
        else
            throw new ArgumentException("Either range or clearPrintArea must be provided");

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);
        return await Task.FromResult(clearPrintArea
            ? $"Print area cleared for sheet {sheetIndex}: {outputPath}"
            : $"Print area set to {range} for sheet {sheetIndex}: {outputPath}");
    }

    /// <summary>
    ///     Sets print titles (rows/columns to repeat on each page)
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional rowsToRepeatAtTop, columnsToRepeatAtLeft</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> SetPrintTitlesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var rows = ArgumentHelper.GetStringNullable(arguments, "rows");
        var columns = ArgumentHelper.GetStringNullable(arguments, "columns");
        var clearTitles = ArgumentHelper.GetBool(arguments, "clearTitles", false);

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (clearTitles)
        {
            worksheet.PageSetup.PrintTitleRows = "";
            worksheet.PageSetup.PrintTitleColumns = "";
        }
        else
        {
            if (!string.IsNullOrEmpty(rows)) worksheet.PageSetup.PrintTitleRows = rows;
            if (!string.IsNullOrEmpty(columns)) worksheet.PageSetup.PrintTitleColumns = columns;
        }

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);
        return await Task.FromResult($"Print titles updated for sheet {sheetIndex}: {outputPath}");
    }

    /// <summary>
    ///     Sets page setup options
    /// </summary>
    /// <param name="arguments">JSON arguments containing various page setup properties</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> SetPageSetupAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var orientation = ArgumentHelper.GetStringNullable(arguments, "orientation");
        var paperSize = ArgumentHelper.GetStringNullable(arguments, "paperSize");
        var leftMargin = ArgumentHelper.GetDoubleNullable(arguments, "leftMargin");
        var rightMargin = ArgumentHelper.GetDoubleNullable(arguments, "rightMargin");
        var topMargin = ArgumentHelper.GetDoubleNullable(arguments, "topMargin");
        var bottomMargin = ArgumentHelper.GetDoubleNullable(arguments, "bottomMargin");
        var header = ArgumentHelper.GetStringNullable(arguments, "header");
        var footer = ArgumentHelper.GetStringNullable(arguments, "footer");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var pageSetup = worksheet.PageSetup;

        if (!string.IsNullOrEmpty(orientation))
            pageSetup.Orientation =
                orientation == "Landscape" ? PageOrientationType.Landscape : PageOrientationType.Portrait;

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

        if (!string.IsNullOrEmpty(header)) pageSetup.SetHeader(0, header);

        if (!string.IsNullOrEmpty(footer)) pageSetup.SetFooter(0, footer);

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);

        var changes = new List<string>();
        if (!string.IsNullOrEmpty(orientation)) changes.Add($"Orientation: {orientation}");
        if (!string.IsNullOrEmpty(paperSize)) changes.Add($"Paper size: {paperSize}");
        if (leftMargin.HasValue || rightMargin.HasValue || topMargin.HasValue || bottomMargin.HasValue)
            changes.Add("Margins set");
        if (!string.IsNullOrEmpty(header)) changes.Add("Header set");
        if (!string.IsNullOrEmpty(footer)) changes.Add("Footer set");

        return await Task.FromResult($"Page setup updated: {string.Join(", ", changes)}");
    }

    /// <summary>
    ///     Sets all print settings at once
    /// </summary>
    /// <param name="arguments">JSON arguments containing all print settings</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> SetAllAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var printArea = ArgumentHelper.GetStringNullable(arguments, "range");
        var printTitleRows = ArgumentHelper.GetStringNullable(arguments, "rows");
        var printTitleColumns = ArgumentHelper.GetStringNullable(arguments, "columns");
        var fitToPage = ArgumentHelper.GetBoolNullable(arguments, "fitToPage");
        var fitToPagesWide = ArgumentHelper.GetIntNullable(arguments, "fitToPagesWide");
        var fitToPagesTall = ArgumentHelper.GetIntNullable(arguments, "fitToPagesTall");
        var orientation = ArgumentHelper.GetStringNullable(arguments, "orientation");
        var paperSize = ArgumentHelper.GetStringNullable(arguments, "paperSize");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pageSetup = worksheet.PageSetup;

        if (!string.IsNullOrEmpty(printArea)) pageSetup.PrintArea = printArea;

        if (!string.IsNullOrEmpty(printTitleRows)) pageSetup.PrintTitleRows = printTitleRows;

        if (!string.IsNullOrEmpty(printTitleColumns)) pageSetup.PrintTitleColumns = printTitleColumns;

        if (fitToPage.HasValue)
        {
            pageSetup.FitToPagesWide = fitToPagesWide ?? 1;
            pageSetup.FitToPagesTall = fitToPagesTall ?? 1;
        }

        if (!string.IsNullOrEmpty(orientation))
            pageSetup.Orientation = orientation.ToLower() == "landscape"
                ? PageOrientationType.Landscape
                : PageOrientationType.Portrait;

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

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);
        return await Task.FromResult($"Print settings updated for sheet {sheetIndex}: {outputPath}");
    }
}