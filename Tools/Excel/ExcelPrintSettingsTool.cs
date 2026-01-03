using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel print settings (print area, titles, page setup, etc.).
///     Merges: ExcelSetPrintAreaTool, ExcelSetPrintTitlesTool, ExcelSetPrintSettingsTool, ExcelSetPageSetupTool.
/// </summary>
[McpServerToolType]
public class ExcelPrintSettingsTool
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
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelPrintSettingsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelPrintSettingsTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_print_settings")]
    [Description(
        @"Manage Excel print settings. Supports 4 operations: set_print_area, set_print_titles, set_page_setup, set_all.

Usage examples:
- Set print area: excel_print_settings(operation='set_print_area', path='book.xlsx', range='A1:D10')
- Set multiple print areas: excel_print_settings(operation='set_print_area', path='book.xlsx', range='A1:D10,F1:H10')
- Set print titles: excel_print_settings(operation='set_print_titles', path='book.xlsx', rows='1:1', columns='A:A')
- Set page setup: excel_print_settings(operation='set_page_setup', path='book.xlsx', orientation='Landscape', paperSize='A4')
- Set margins: excel_print_settings(operation='set_page_setup', path='book.xlsx', leftMargin=0.5, topMargin=0.75)
- Set fit to page: excel_print_settings(operation='set_all', path='book.xlsx', fitToPage=true, fitToPagesWide=1, fitToPagesTall=0)
- Set all: excel_print_settings(operation='set_all', path='book.xlsx', range='A1:D10', orientation='Portrait')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'set_print_area': Set print area (required params: path, range or clearPrintArea)
- 'set_print_titles': Set print titles (required params: path)
- 'set_page_setup': Set page setup (required params: path)
- 'set_all': Set all print settings (required params: path)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description(
            "Print area range. Supports single range (e.g., 'A1:D10') or multiple ranges separated by comma (e.g., 'A1:D10,F1:H10')")]
        string? range = null,
        [Description("Clear print area (optional, for set_print_area, default: false)")]
        bool clearPrintArea = false,
        [Description("Rows to repeat on each printed page (e.g., '1:1' for first row, '1:2' for first two rows)")]
        string? rows = null,
        [Description(
            "Columns to repeat on each printed page (e.g., 'A:A' for first column, 'A:B' for first two columns)")]
        string? columns = null,
        [Description("Clear print titles (optional, for set_print_titles, default: false)")]
        bool clearTitles = false,
        [Description("Page orientation (optional, default: Portrait)")]
        string? orientation = null,
        [Description(
            "Paper size. Supported values: A3, A4, A5, B4, B5, Letter, Legal, Tabloid, Executive (optional, default: A4)")]
        string? paperSize = null,
        [Description("Left margin in inches (optional, default: 0.7)")]
        double? leftMargin = null,
        [Description("Right margin in inches (optional, default: 0.7)")]
        double? rightMargin = null,
        [Description("Top margin in inches (optional, default: 0.75)")]
        double? topMargin = null,
        [Description("Bottom margin in inches (optional, default: 0.75)")]
        double? bottomMargin = null,
        [Description("Header text for center section (optional)")]
        string? header = null,
        [Description("Footer text for center section (optional)")]
        string? footer = null,
        [Description("Enable fit to page mode. When true, fitToPagesWide and fitToPagesTall are used.")]
        bool? fitToPage = null,
        [Description(
            "Number of pages wide to fit content (optional, default: 1 when fitToPage is true, use 0 for automatic)")]
        int? fitToPagesWide = null,
        [Description(
            "Number of pages tall to fit content (optional, default: 1 when fitToPage is true, use 0 for automatic)")]
        int? fitToPagesTall = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "set_print_area" => SetPrintArea(ctx, outputPath, sheetIndex, range, clearPrintArea),
            "set_print_titles" => SetPrintTitles(ctx, outputPath, sheetIndex, rows, columns, clearTitles),
            "set_page_setup" => SetPageSetup(ctx, outputPath, sheetIndex, orientation, paperSize, leftMargin,
                rightMargin, topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall),
            "set_all" => SetAll(ctx, outputPath, sheetIndex, range, rows, columns, orientation, paperSize, leftMargin,
                rightMargin, topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets print area for the worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The print area range (e.g., 'A1:D10').</param>
    /// <param name="clearPrintArea">Whether to clear the print area.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when neither range nor clearPrintArea is provided.</exception>
    private static string SetPrintArea(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? range, bool clearPrintArea)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (clearPrintArea)
            worksheet.PageSetup.PrintArea = "";
        else if (!string.IsNullOrEmpty(range))
            worksheet.PageSetup.PrintArea = range;
        else
            throw new ArgumentException("Either range or clearPrintArea must be provided");

        ctx.Save(outputPath);
        return clearPrintArea
            ? $"Print area cleared for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}"
            : $"Print area set to {range} for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets print titles (rows/columns to repeat on each page).
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="rows">The rows to repeat on each page (e.g., '1:1').</param>
    /// <param name="columns">The columns to repeat on each page (e.g., 'A:A').</param>
    /// <param name="clearTitles">Whether to clear the print titles.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetPrintTitles(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? rows, string? columns, bool clearTitles)
    {
        var workbook = ctx.Document;
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

        ctx.Save(outputPath);
        return $"Print titles updated for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets page setup options.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="orientation">The page orientation ('Portrait' or 'Landscape').</param>
    /// <param name="paperSize">The paper size (e.g., 'A4', 'Letter').</param>
    /// <param name="leftMargin">The left margin in inches.</param>
    /// <param name="rightMargin">The right margin in inches.</param>
    /// <param name="topMargin">The top margin in inches.</param>
    /// <param name="bottomMargin">The bottom margin in inches.</param>
    /// <param name="header">The header text for center section.</param>
    /// <param name="footer">The footer text for center section.</param>
    /// <param name="fitToPage">Whether to enable fit to page mode.</param>
    /// <param name="fitToPagesWide">The number of pages wide to fit content.</param>
    /// <param name="fitToPagesTall">The number of pages tall to fit content.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetPageSetup(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? orientation, string? paperSize, double? leftMargin, double? rightMargin,
        double? topMargin, double? bottomMargin, string? header, string? footer,
        bool? fitToPage, int? fitToPagesWide, int? fitToPagesTall)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pageSetup = worksheet.PageSetup;

        var changes = ApplyPageSetup(pageSetup, orientation, paperSize, leftMargin, rightMargin,
            topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall);

        ctx.Save(outputPath);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return $"Page setup updated ({changesStr}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Applies page setup options to the PageSetup object.
    /// </summary>
    /// <param name="pageSetup">The PageSetup object to modify.</param>
    /// <param name="orientation">The page orientation ('Portrait' or 'Landscape').</param>
    /// <param name="paperSize">The paper size (e.g., 'A4', 'Letter').</param>
    /// <param name="leftMargin">The left margin in inches.</param>
    /// <param name="rightMargin">The right margin in inches.</param>
    /// <param name="topMargin">The top margin in inches.</param>
    /// <param name="bottomMargin">The bottom margin in inches.</param>
    /// <param name="header">The header text for center section.</param>
    /// <param name="footer">The footer text for center section.</param>
    /// <param name="fitToPage">Whether to enable fit to page mode.</param>
    /// <param name="fitToPagesWide">The number of pages wide to fit content.</param>
    /// <param name="fitToPagesTall">The number of pages tall to fit content.</param>
    /// <returns>A list of change descriptions indicating what settings were modified.</returns>
    /// <exception cref="ArgumentException">Thrown when an invalid paper size is specified.</exception>
    private static List<string> ApplyPageSetup(PageSetup pageSetup, string? orientation, string? paperSize,
        double? leftMargin, double? rightMargin, double? topMargin, double? bottomMargin,
        string? header, string? footer, bool? fitToPage, int? fitToPagesWide, int? fitToPagesTall)
    {
        List<string> changes = [];

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
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="printArea">The print area range (e.g., 'A1:D10').</param>
    /// <param name="printTitleRows">The rows to repeat on each page (e.g., '1:1').</param>
    /// <param name="printTitleColumns">The columns to repeat on each page (e.g., 'A:A').</param>
    /// <param name="orientation">The page orientation ('Portrait' or 'Landscape').</param>
    /// <param name="paperSize">The paper size (e.g., 'A4', 'Letter').</param>
    /// <param name="leftMargin">The left margin in inches.</param>
    /// <param name="rightMargin">The right margin in inches.</param>
    /// <param name="topMargin">The top margin in inches.</param>
    /// <param name="bottomMargin">The bottom margin in inches.</param>
    /// <param name="header">The header text for center section.</param>
    /// <param name="footer">The footer text for center section.</param>
    /// <param name="fitToPage">Whether to enable fit to page mode.</param>
    /// <param name="fitToPagesWide">The number of pages wide to fit content.</param>
    /// <param name="fitToPagesTall">The number of pages tall to fit content.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetAll(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? printArea, string? printTitleRows, string? printTitleColumns,
        string? orientation, string? paperSize, double? leftMargin, double? rightMargin,
        double? topMargin, double? bottomMargin, string? header, string? footer,
        bool? fitToPage, int? fitToPagesWide, int? fitToPagesTall)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pageSetup = worksheet.PageSetup;

        List<string> changes = [];

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

        var pageSetupChanges = ApplyPageSetup(pageSetup, orientation, paperSize, leftMargin, rightMargin,
            topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall);
        changes.AddRange(pageSetupChanges);

        ctx.Save(outputPath);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return $"Print settings updated for sheet {sheetIndex} ({changesStr}). {ctx.GetOutputMessage(outputPath)}";
    }
}