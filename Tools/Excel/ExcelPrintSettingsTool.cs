using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel print settings (print area, titles, page setup, etc.).
///     Merges: ExcelSetPrintAreaTool, ExcelSetPrintTitlesTool, ExcelSetPrintSettingsTool, ExcelSetPageSetupTool.
/// </summary>
public class ExcelPrintSettingsTool : IAsposeTool
{
    /// <summary>
    ///     Supported paper size mappings.
    /// </summary>
    private static readonly Dictionary<string, PaperSizeType> PaperSizeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["A3"] = PaperSizeType.PaperA3,
        ["A4"] = PaperSizeType.PaperA4,
        ["A5"] = PaperSizeType.PaperA5,
        ["B4"] = PaperSizeType.PaperB4,
        ["B5"] = PaperSizeType.PaperB5,
        ["Letter"] = PaperSizeType.PaperLetter,
        ["Legal"] = PaperSizeType.PaperLegal,
        ["Tabloid"] = PaperSizeType.PaperTabloid,
        ["Executive"] = PaperSizeType.PaperExecutive
    };

    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description =>
        @"Manage Excel print settings. Supports 4 operations: set_print_area, set_print_titles, set_page_setup, set_all.

Usage examples:
- Set print area: excel_print_settings(operation='set_print_area', path='book.xlsx', range='A1:D10')
- Set multiple print areas: excel_print_settings(operation='set_print_area', path='book.xlsx', range='A1:D10,F1:H10')
- Set print titles: excel_print_settings(operation='set_print_titles', path='book.xlsx', rows='1:1', columns='A:A')
- Set page setup: excel_print_settings(operation='set_page_setup', path='book.xlsx', orientation='Landscape', paperSize='A4')
- Set margins: excel_print_settings(operation='set_page_setup', path='book.xlsx', leftMargin=0.5, topMargin=0.75)
- Set fit to page: excel_print_settings(operation='set_all', path='book.xlsx', fitToPage=true, fitToPagesWide=1, fitToPagesTall=0)
- Set all: excel_print_settings(operation='set_all', path='book.xlsx', range='A1:D10', orientation='Portrait')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool.
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
- 'set_print_area': Set print area (required params: path, range or clearPrintArea)
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
                description =
                    "Print area range. Supports single range (e.g., 'A1:D10') or multiple ranges separated by comma (e.g., 'A1:D10,F1:H10'). Required for set_print_area unless clearPrintArea is true."
            },
            clearPrintArea = new
            {
                type = "boolean",
                description = "Clear print area (optional, for set_print_area, default: false)"
            },
            rows = new
            {
                type = "string",
                description =
                    "Rows to repeat on each printed page (e.g., '1:1' for first row, '1:2' for first two rows). Optional for set_print_titles."
            },
            columns = new
            {
                type = "string",
                description =
                    "Columns to repeat on each printed page (e.g., 'A:A' for first column, 'A:B' for first two columns). Optional for set_print_titles."
            },
            clearTitles = new
            {
                type = "boolean",
                description = "Clear print titles (optional, for set_print_titles, default: false)"
            },
            orientation = new
            {
                type = "string",
                description = "Page orientation (optional, default: Portrait)",
                @enum = new[] { "Portrait", "Landscape" }
            },
            paperSize = new
            {
                type = "string",
                description =
                    "Paper size. Supported values: A3, A4, A5, B4, B5, Letter, Legal, Tabloid, Executive (optional, default: A4)",
                @enum = new[] { "A3", "A4", "A5", "B4", "B5", "Letter", "Legal", "Tabloid", "Executive" }
            },
            leftMargin = new
            {
                type = "number",
                description = "Left margin in inches (optional, default: 0.7)"
            },
            rightMargin = new
            {
                type = "number",
                description = "Right margin in inches (optional, default: 0.7)"
            },
            topMargin = new
            {
                type = "number",
                description = "Top margin in inches (optional, default: 0.75)"
            },
            bottomMargin = new
            {
                type = "number",
                description = "Bottom margin in inches (optional, default: 0.75)"
            },
            header = new
            {
                type = "string",
                description = "Header text for center section (optional)"
            },
            footer = new
            {
                type = "string",
                description = "Footer text for center section (optional)"
            },
            fitToPage = new
            {
                type = "boolean",
                description =
                    "Enable fit to page mode. When true, fitToPagesWide and fitToPagesTall are used. Note: This disables percentage zoom scaling."
            },
            fitToPagesWide = new
            {
                type = "number",
                description =
                    "Number of pages wide to fit content (optional, default: 1 when fitToPage is true, use 0 for automatic)"
            },
            fitToPagesTall = new
            {
                type = "number",
                description =
                    "Number of pages tall to fit content (optional, default: 1 when fitToPage is true, use 0 for automatic)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or parameters are invalid.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "set_print_area" => await SetPrintAreaAsync(path, outputPath, sheetIndex, arguments),
            "set_print_titles" => await SetPrintTitlesAsync(path, outputPath, sheetIndex, arguments),
            "set_page_setup" => await SetPageSetupAsync(path, outputPath, sheetIndex, arguments),
            "set_all" => await SetAllAsync(path, outputPath, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets print area for the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing optional range, clearPrintArea.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when neither range nor clearPrintArea is provided.</exception>
    private Task<string> SetPrintAreaAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
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

            workbook.Save(outputPath);
            return clearPrintArea
                ? $"Print area cleared for sheet {sheetIndex}. Output: {outputPath}"
                : $"Print area set to {range} for sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets print titles (rows/columns to repeat on each page).
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing optional rows, columns, clearTitles.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when sheet index is out of range.</exception>
    private Task<string> SetPrintTitlesAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
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

            workbook.Save(outputPath);
            return $"Print titles updated for sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets page setup options.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing various page setup properties.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when sheet index is out of range or paper size is invalid.</exception>
    private Task<string> SetPageSetupAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pageSetup = worksheet.PageSetup;

            var changes = ApplyPageSetup(pageSetup, arguments);

            workbook.Save(outputPath);

            var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
            return $"Page setup updated ({changesStr}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Applies page setup options from arguments to the PageSetup object.
    /// </summary>
    /// <param name="pageSetup">The PageSetup object to configure.</param>
    /// <param name="arguments">JSON arguments containing page setup properties.</param>
    /// <returns>List of changes applied.</returns>
    /// <exception cref="ArgumentException">Thrown when paper size is invalid.</exception>
    private static List<string> ApplyPageSetup(PageSetup pageSetup, JsonObject? arguments)
    {
        var changes = new List<string>();

        var orientation = ArgumentHelper.GetStringNullable(arguments, "orientation");
        var paperSize = ArgumentHelper.GetStringNullable(arguments, "paperSize");
        var leftMargin = ArgumentHelper.GetDoubleNullable(arguments, "leftMargin");
        var rightMargin = ArgumentHelper.GetDoubleNullable(arguments, "rightMargin");
        var topMargin = ArgumentHelper.GetDoubleNullable(arguments, "topMargin");
        var bottomMargin = ArgumentHelper.GetDoubleNullable(arguments, "bottomMargin");
        var header = ArgumentHelper.GetStringNullable(arguments, "header");
        var footer = ArgumentHelper.GetStringNullable(arguments, "footer");
        var fitToPage = ArgumentHelper.GetBoolNullable(arguments, "fitToPage");
        var fitToPagesWide = ArgumentHelper.GetIntNullable(arguments, "fitToPagesWide");
        var fitToPagesTall = ArgumentHelper.GetIntNullable(arguments, "fitToPagesTall");

        if (!string.IsNullOrEmpty(orientation))
        {
            pageSetup.Orientation = string.Equals(orientation, "Landscape", StringComparison.OrdinalIgnoreCase)
                ? PageOrientationType.Landscape
                : PageOrientationType.Portrait;
            changes.Add($"orientation={orientation}");
        }

        if (!string.IsNullOrEmpty(paperSize))
        {
            if (PaperSizeMap.TryGetValue(paperSize, out var size))
            {
                pageSetup.PaperSize = size;
                changes.Add($"paperSize={paperSize}");
            }
            else
            {
                throw new ArgumentException(
                    $"Invalid paper size: '{paperSize}'. Supported values: {string.Join(", ", PaperSizeMap.Keys)}");
            }
        }

        if (leftMargin.HasValue)
        {
            pageSetup.LeftMargin = leftMargin.Value;
            changes.Add($"leftMargin={leftMargin.Value}");
        }

        if (rightMargin.HasValue)
        {
            pageSetup.RightMargin = rightMargin.Value;
            changes.Add($"rightMargin={rightMargin.Value}");
        }

        if (topMargin.HasValue)
        {
            pageSetup.TopMargin = topMargin.Value;
            changes.Add($"topMargin={topMargin.Value}");
        }

        if (bottomMargin.HasValue)
        {
            pageSetup.BottomMargin = bottomMargin.Value;
            changes.Add($"bottomMargin={bottomMargin.Value}");
        }

        if (!string.IsNullOrEmpty(header))
        {
            pageSetup.SetHeader(1, header);
            changes.Add("header");
        }

        if (!string.IsNullOrEmpty(footer))
        {
            pageSetup.SetFooter(1, footer);
            changes.Add("footer");
        }

        if (fitToPage == true)
        {
            pageSetup.FitToPagesWide = fitToPagesWide ?? 1;
            pageSetup.FitToPagesTall = fitToPagesTall ?? 1;
            changes.Add($"fitToPage(wide={pageSetup.FitToPagesWide}, tall={pageSetup.FitToPagesTall})");
        }

        return changes;
    }

    /// <summary>
    ///     Sets all print settings at once.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing all print settings.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when sheet index is out of range or paper size is invalid.</exception>
    private Task<string> SetAllAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var printArea = ArgumentHelper.GetStringNullable(arguments, "range");
            var printTitleRows = ArgumentHelper.GetStringNullable(arguments, "rows");
            var printTitleColumns = ArgumentHelper.GetStringNullable(arguments, "columns");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pageSetup = worksheet.PageSetup;

            var changes = new List<string>();

            if (!string.IsNullOrEmpty(printArea))
            {
                pageSetup.PrintArea = printArea;
                changes.Add($"printArea={printArea}");
            }

            if (!string.IsNullOrEmpty(printTitleRows))
            {
                pageSetup.PrintTitleRows = printTitleRows;
                changes.Add($"printTitleRows={printTitleRows}");
            }

            if (!string.IsNullOrEmpty(printTitleColumns))
            {
                pageSetup.PrintTitleColumns = printTitleColumns;
                changes.Add($"printTitleColumns={printTitleColumns}");
            }

            var pageSetupChanges = ApplyPageSetup(pageSetup, arguments);
            changes.AddRange(pageSetupChanges);

            workbook.Save(outputPath);

            var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
            return $"Print settings updated for sheet {sheetIndex} ({changesStr}). Output: {outputPath}";
        });
    }
}