using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Properties;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel properties (workbook properties, sheet properties, sheet info).
///     Merges: ExcelGetWorkbookPropertiesTool, ExcelSetWorkbookPropertiesTool, ExcelGetSheetPropertiesTool,
///     ExcelEditSheetPropertiesTool, ExcelGetSheetInfoTool.
/// </summary>
[McpServerToolType]
public class ExcelPropertiesTool
{
    /// <summary>
    ///     Operation name for getting workbook properties.
    /// </summary>
    private const string OperationGetWorkbookProperties = "get_workbook_properties";

    /// <summary>
    ///     Operation name for setting workbook properties.
    /// </summary>
    private const string OperationSetWorkbookProperties = "set_workbook_properties";

    /// <summary>
    ///     Operation name for getting sheet properties.
    /// </summary>
    private const string OperationGetSheetProperties = "get_sheet_properties";

    /// <summary>
    ///     Operation name for editing sheet properties.
    /// </summary>
    private const string OperationEditSheetProperties = "edit_sheet_properties";

    /// <summary>
    ///     Operation name for getting sheet info.
    /// </summary>
    private const string OperationGetSheetInfo = "get_sheet_info";

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelPropertiesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelPropertiesTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_properties")]
    [Description(
        @"Manage Excel properties. Supports 5 operations: get_workbook_properties, set_workbook_properties, get_sheet_properties, edit_sheet_properties, get_sheet_info.

Usage examples:
- Get workbook properties: excel_properties(operation='get_workbook_properties', path='book.xlsx')
- Set workbook properties: excel_properties(operation='set_workbook_properties', path='book.xlsx', title='Title', author='Author')
- Get sheet properties: excel_properties(operation='get_sheet_properties', path='book.xlsx', sheetIndex=0)
- Edit sheet properties: excel_properties(operation='edit_sheet_properties', path='book.xlsx', sheetIndex=0, name='New Name')
- Get sheet info: excel_properties(operation='get_sheet_info', path='book.xlsx')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'get_workbook_properties': Get workbook properties (required params: path)
- 'set_workbook_properties': Set workbook properties (required params: path)
- 'get_sheet_properties': Get sheet properties (required params: path, sheetIndex)
- 'edit_sheet_properties': Edit sheet properties (required params: path, sheetIndex)
- 'get_sheet_info': Get sheet info (required params: path)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0, required for sheet operations)")]
        int sheetIndex = 0,
        [Description("Title (optional, for set_workbook_properties)")]
        string? title = null,
        [Description("Subject (optional, for set_workbook_properties)")]
        string? subject = null,
        [Description("Author (optional, for set_workbook_properties)")]
        string? author = null,
        [Description("Keywords (optional, for set_workbook_properties)")]
        string? keywords = null,
        [Description("Comments (optional, for set_workbook_properties)")]
        string? comments = null,
        [Description("Category (optional, for set_workbook_properties)")]
        string? category = null,
        [Description("Company (optional, for set_workbook_properties)")]
        string? company = null,
        [Description("Manager (optional, for set_workbook_properties)")]
        string? manager = null,
        [Description("Custom properties as JSON object (optional, for set_workbook_properties)")]
        string? customProperties = null,
        [Description("Sheet name (optional, for edit_sheet_properties)")]
        string? name = null,
        [Description("Sheet visibility (optional, for edit_sheet_properties)")]
        bool? isVisible = null,
        [Description("Tab color hex (e.g., #FF0000, optional, for edit_sheet_properties)")]
        string? tabColor = null,
        [Description("Set as selected sheet (optional, for edit_sheet_properties)")]
        bool? isSelected = null,
        [Description("Sheet index for get_sheet_info (optional, if not provided returns all sheets)")]
        int? targetSheetIndex = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            OperationGetWorkbookProperties => GetWorkbookProperties(ctx),
            OperationSetWorkbookProperties => SetWorkbookProperties(ctx, outputPath, title, subject, author, keywords,
                comments, category, company, manager, customProperties),
            OperationGetSheetProperties => GetSheetProperties(ctx, sheetIndex),
            OperationEditSheetProperties => EditSheetProperties(ctx, outputPath, sheetIndex, name, isVisible, tabColor,
                isSelected),
            OperationGetSheetInfo => GetSheetInfo(ctx, targetSheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets workbook properties.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <returns>A JSON string containing the workbook properties.</returns>
    private static string GetWorkbookProperties(DocumentContext<Workbook> ctx)
    {
        var workbook = ctx.Document;
        var props = workbook.BuiltInDocumentProperties;
        var customProps = workbook.CustomDocumentProperties;

        List<object> customPropsList = [];
        if (customProps.Count > 0)
            foreach (var prop in customProps)
                customPropsList.Add(new
                    { name = prop.Name, value = prop.Value?.ToString(), type = prop.Type.ToString() });

        var result = new
        {
            title = props.Title,
            subject = props.Subject,
            author = props.Author,
            keywords = props.Keywords,
            comments = props.Comments,
            category = props.Category,
            company = props.Company,
            manager = props.Manager,
            created = props.CreatedTime.ToString("o"),
            modified = props.LastSavedTime.ToString("o"),
            lastSavedBy = props.LastSavedBy,
            revision = props.RevisionNumber,
            totalSheets = workbook.Worksheets.Count,
            customProperties = customPropsList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Sets workbook properties.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="title">The document title.</param>
    /// <param name="subject">The document subject.</param>
    /// <param name="author">The document author.</param>
    /// <param name="keywords">The document keywords.</param>
    /// <param name="comments">The document comments.</param>
    /// <param name="category">The document category.</param>
    /// <param name="company">The company name.</param>
    /// <param name="manager">The manager name.</param>
    /// <param name="customPropertiesJson">Custom properties as a JSON object.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when customPropertiesJson has an invalid JSON format.</exception>
    private static string SetWorkbookProperties(DocumentContext<Workbook> ctx, string? outputPath,
        string? title, string? subject, string? author, string? keywords, string? comments,
        string? category, string? company, string? manager, string? customPropertiesJson)
    {
        var workbook = ctx.Document;
        var props = workbook.BuiltInDocumentProperties;

        if (!string.IsNullOrEmpty(title)) props.Title = title;
        if (!string.IsNullOrEmpty(subject)) props.Subject = subject;
        if (!string.IsNullOrEmpty(author)) props.Author = author;
        if (!string.IsNullOrEmpty(keywords)) props.Keywords = keywords;
        if (!string.IsNullOrEmpty(comments)) props.Comments = comments;
        if (!string.IsNullOrEmpty(category)) props.Category = category;
        if (!string.IsNullOrEmpty(company)) props.Company = company;
        if (!string.IsNullOrEmpty(manager)) props.Manager = manager;

        if (!string.IsNullOrEmpty(customPropertiesJson))
            try
            {
                var customProps = JsonNode.Parse(customPropertiesJson)?.AsObject();
                if (customProps != null)
                    foreach (var kvp in customProps)
                    {
                        var value = kvp.Value?.GetValue<string>() ?? "";
                        var existingProp = FindCustomProperty(workbook.CustomDocumentProperties, kvp.Key);
                        if (existingProp != null)
                            workbook.CustomDocumentProperties.Remove(kvp.Key);
                        workbook.CustomDocumentProperties.Add(kvp.Key, value);
                    }
            }
            catch (JsonException ex)
            {
                throw new ArgumentException($"Invalid JSON format for customProperties: {ex.Message}");
            }

        ctx.Save(outputPath);
        return $"Workbook properties updated successfully. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Finds a custom property by name.
    /// </summary>
    /// <param name="customProperties">The collection of custom document properties.</param>
    /// <param name="name">The name of the property to find.</param>
    /// <returns>The document property if found; otherwise, null.</returns>
    private static DocumentProperty? FindCustomProperty(CustomDocumentPropertyCollection customProperties, string name)
    {
        foreach (var prop in customProperties)
            if (string.Equals(prop.Name, name, StringComparison.OrdinalIgnoreCase))
                return prop;
        return null;
    }

    /// <summary>
    ///     Gets worksheet properties.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The zero-based index of the worksheet.</param>
    /// <returns>A JSON string containing the worksheet properties.</returns>
    private static string GetSheetProperties(DocumentContext<Workbook> ctx, int sheetIndex)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pageSetup = worksheet.PageSetup;

        var result = new
        {
            name = worksheet.Name,
            index = sheetIndex,
            isVisible = worksheet.IsVisible,
            tabColor = worksheet.TabColor.ToString(),
            isSelected = workbook.Worksheets.ActiveSheetIndex == sheetIndex,
            dataRowCount = worksheet.Cells.MaxDataRow + 1,
            dataColumnCount = worksheet.Cells.MaxDataColumn + 1,
            isProtected = worksheet.Protection.IsProtectedWithPassword,
            commentsCount = worksheet.Comments.Count,
            chartsCount = worksheet.Charts.Count,
            picturesCount = worksheet.Pictures.Count,
            hyperlinksCount = worksheet.Hyperlinks.Count,
            printSettings = new
            {
                printArea = pageSetup.PrintArea,
                printTitleRows = pageSetup.PrintTitleRows,
                printTitleColumns = pageSetup.PrintTitleColumns,
                orientation = pageSetup.Orientation.ToString(),
                paperSize = pageSetup.PaperSize.ToString(),
                fitToPagesWide = pageSetup.FitToPagesWide,
                fitToPagesTall = pageSetup.FitToPagesTall
            }
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Edits worksheet properties.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The zero-based index of the worksheet.</param>
    /// <param name="name">The new name for the worksheet.</param>
    /// <param name="isVisible">Whether the worksheet should be visible.</param>
    /// <param name="tabColor">The tab color in hex format (e.g., '#FF0000').</param>
    /// <param name="isSelected">Whether to set this worksheet as the selected sheet.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string EditSheetProperties(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? name, bool? isVisible, string? tabColor, bool? isSelected)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (!string.IsNullOrEmpty(name)) worksheet.Name = name;

        if (isVisible.HasValue) worksheet.IsVisible = isVisible.Value;

        if (!string.IsNullOrWhiteSpace(tabColor))
        {
            var color = ColorHelper.ParseColor(tabColor);
            worksheet.TabColor = color;
        }

        if (isSelected.HasValue && isSelected.Value) workbook.Worksheets.ActiveSheetIndex = sheetIndex;

        ctx.Save(outputPath);
        return $"Sheet {sheetIndex} properties updated successfully. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets information about all worksheets or a specific worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The zero-based index of the worksheet, or null to get all worksheets.</param>
    /// <returns>A JSON string containing the worksheet information.</returns>
    private static string GetSheetInfo(DocumentContext<Workbook> ctx, int? sheetIndex)
    {
        var workbook = ctx.Document;

        List<object> sheetList = [];

        if (sheetIndex.HasValue)
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);
            sheetList.Add(CreateSheetInfo(worksheet, sheetIndex.Value));
        }
        else
        {
            for (var i = 0; i < workbook.Worksheets.Count; i++)
                sheetList.Add(CreateSheetInfo(workbook.Worksheets[i], i));
        }

        var result = new
        {
            count = sheetList.Count,
            totalWorksheets = workbook.Worksheets.Count,
            items = sheetList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Creates sheet information object for JSON serialization.
    /// </summary>
    /// <param name="worksheet">The worksheet to extract information from.</param>
    /// <param name="index">The zero-based index of the worksheet.</param>
    /// <returns>An anonymous object containing the worksheet information.</returns>
    private static object CreateSheetInfo(Worksheet worksheet, int index)
    {
        return new
        {
            index,
            name = worksheet.Name,
            visibility = worksheet.VisibilityType.ToString(),
            dataRowCount = worksheet.Cells.MaxDataRow + 1,
            dataColumnCount = worksheet.Cells.MaxDataColumn + 1,
            usedRange = new
            {
                rowCount = worksheet.Cells.MaxRow + 1,
                columnCount = worksheet.Cells.MaxColumn + 1
            },
            pageOrientation = worksheet.PageSetup.Orientation.ToString(),
            paperSize = worksheet.PageSetup.PaperSize.ToString(),
            freezePanes = new
            {
                row = worksheet.FirstVisibleRow,
                column = worksheet.FirstVisibleColumn
            }
        };
    }
}