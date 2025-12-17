using System.Text;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel properties (workbook properties, sheet properties, sheet info)
///     Merges: ExcelGetWorkbookPropertiesTool, ExcelSetWorkbookPropertiesTool, ExcelGetSheetPropertiesTool,
///     ExcelEditSheetPropertiesTool, ExcelGetSheetInfoTool
/// </summary>
public class ExcelPropertiesTool : IAsposeTool
{
    /// <summary>    ///     Gets the description of the tool and its usage examples    /// </summary>    public string Description =>
        @"Manage Excel properties. Supports 5 operations: get_workbook_properties, set_workbook_properties, get_sheet_properties, edit_sheet_properties, get_sheet_info.

Usage examples:
- Get workbook properties: excel_properties(operation='get_workbook_properties', path='book.xlsx')
- Set workbook properties: excel_properties(operation='set_workbook_properties', path='book.xlsx', title='Title', author='Author')
- Get sheet properties: excel_properties(operation='get_sheet_properties', path='book.xlsx', sheetIndex=0)
- Edit sheet properties: excel_properties(operation='edit_sheet_properties', path='book.xlsx', sheetIndex=0, name='New Name')
- Get sheet info: excel_properties(operation='get_sheet_info', path='book.xlsx')";

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
- 'get_workbook_properties': Get workbook properties (required params: path)
- 'set_workbook_properties': Set workbook properties (required params: path)
- 'get_sheet_properties': Get sheet properties (required params: path, sheetIndex)
- 'edit_sheet_properties': Edit sheet properties (required params: path, sheetIndex)
- 'get_sheet_info': Get sheet info (required params: path)",
                @enum = new[]
                {
                    "get_workbook_properties", "set_workbook_properties", "edit_sheet_properties",
                    "get_sheet_properties", "get_sheet_info"
                }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0, required for sheet operations)"
            },
            title = new
            {
                type = "string",
                description = "Title (optional, for set_workbook_properties)"
            },
            subject = new
            {
                type = "string",
                description = "Subject (optional, for set_workbook_properties)"
            },
            author = new
            {
                type = "string",
                description = "Author (optional, for set_workbook_properties)"
            },
            keywords = new
            {
                type = "string",
                description = "Keywords (optional, for set_workbook_properties)"
            },
            comments = new
            {
                type = "string",
                description = "Comments (optional, for set_workbook_properties)"
            },
            category = new
            {
                type = "string",
                description = "Category (optional, for set_workbook_properties)"
            },
            company = new
            {
                type = "string",
                description = "Company (optional, for set_workbook_properties)"
            },
            manager = new
            {
                type = "string",
                description = "Manager (optional, for set_workbook_properties)"
            },
            customProperties = new
            {
                type = "object",
                description = "Custom properties as key-value pairs (optional, for set_workbook_properties)"
            },
            name = new
            {
                type = "string",
                description = "Sheet name (optional, for edit_sheet_properties)"
            },
            isVisible = new
            {
                type = "boolean",
                description = "Sheet visibility (optional, for edit_sheet_properties)"
            },
            tabColor = new
            {
                type = "string",
                description = "Tab color hex (e.g., #FF0000, optional, for edit_sheet_properties)"
            },
            isSelected = new
            {
                type = "boolean",
                description = "Set as selected sheet (optional, for edit_sheet_properties)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for set/edit_sheet_properties operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "get_workbook_properties" => await GetWorkbookPropertiesAsync(arguments, path),
            "set_workbook_properties" => await SetWorkbookPropertiesAsync(arguments, path),
            "get_sheet_properties" => await GetSheetPropertiesAsync(arguments, path, sheetIndex),
            "edit_sheet_properties" => await EditSheetPropertiesAsync(arguments, path, sheetIndex),
            "get_sheet_info" => await GetSheetInfoAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets workbook properties
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Formatted string with workbook properties</returns>
    private async Task<string> GetWorkbookPropertiesAsync(JsonObject? _, string path)
    {
        using var workbook = new Workbook(path);
        var props = workbook.BuiltInDocumentProperties;
        var customProps = workbook.CustomDocumentProperties;

        var sb = new StringBuilder();
        sb.AppendLine("Workbook Properties:");
        sb.AppendLine($"  Title: {props.Title ?? "(none)"}");
        sb.AppendLine($"  Subject: {props.Subject ?? "(none)"}");
        sb.AppendLine($"  Author: {props.Author ?? "(none)"}");
        sb.AppendLine($"  Keywords: {props.Keywords ?? "(none)"}");
        sb.AppendLine($"  Comments: {props.Comments ?? "(none)"}");
        sb.AppendLine($"  Category: {props.Category ?? "(none)"}");
        sb.AppendLine($"  Company: {props.Company ?? "(none)"}");
        sb.AppendLine($"  Manager: {props.Manager ?? "(none)"}");
        sb.AppendLine($"  Created: {props.CreatedTime}");
        sb.AppendLine($"  Modified: {props.LastSavedTime}");
        sb.AppendLine($"  Last Saved By: {props.LastSavedBy ?? "(none)"}");
        sb.AppendLine($"  Revision: {props.RevisionNumber}");

        if (customProps.Count > 0)
        {
            sb.AppendLine("\nCustom Properties:");
            foreach (var prop in customProps) sb.AppendLine($"  {prop.Name}: {prop.Value}");
        }

        sb.AppendLine($"\nTotal Sheets: {workbook.Worksheets.Count}");

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    ///     Sets workbook properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing various property values</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message</returns>
    private async Task<string> SetWorkbookPropertiesAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var title = ArgumentHelper.GetStringNullable(arguments, "title");
        var subject = ArgumentHelper.GetStringNullable(arguments, "subject");
        var author = ArgumentHelper.GetStringNullable(arguments, "author");
        var keywords = ArgumentHelper.GetStringNullable(arguments, "keywords");
        var comments = ArgumentHelper.GetStringNullable(arguments, "comments");
        var category = ArgumentHelper.GetStringNullable(arguments, "category");
        var company = ArgumentHelper.GetStringNullable(arguments, "company");
        var manager = ArgumentHelper.GetStringNullable(arguments, "manager");
        var customProps = ArgumentHelper.GetObject(arguments, "customProperties", false);

        using var workbook = new Workbook(path);
        var props = workbook.BuiltInDocumentProperties;

        if (!string.IsNullOrEmpty(title)) props.Title = title;
        if (!string.IsNullOrEmpty(subject)) props.Subject = subject;
        if (!string.IsNullOrEmpty(author)) props.Author = author;
        if (!string.IsNullOrEmpty(keywords)) props.Keywords = keywords;
        if (!string.IsNullOrEmpty(comments)) props.Comments = comments;
        if (!string.IsNullOrEmpty(category)) props.Category = category;
        if (!string.IsNullOrEmpty(company)) props.Company = company;
        if (!string.IsNullOrEmpty(manager)) props.Manager = manager;

        if (customProps != null)
            foreach (var kvp in customProps)
                workbook.CustomDocumentProperties.Add(kvp.Key, kvp.Value?.GetValue<string>() ?? "");

        workbook.Save(outputPath);
        return await Task.FromResult($"Workbook properties updated: {outputPath}");
    }

    /// <summary>
    ///     Gets worksheet properties
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with worksheet properties</returns>
    private async Task<string> GetSheetPropertiesAsync(JsonObject? _, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var sb = new StringBuilder();

        sb.AppendLine("Sheet Properties:");
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
        sb.AppendLine("\nPrint Settings:");
        sb.AppendLine($"  Print Area: {pageSetup.PrintArea ?? "(none)"}");
        sb.AppendLine($"  Print Title Rows: {pageSetup.PrintTitleRows ?? "(none)"}");
        sb.AppendLine($"  Print Title Columns: {pageSetup.PrintTitleColumns ?? "(none)"}");
        sb.AppendLine($"  Orientation: {pageSetup.Orientation}");
        sb.AppendLine($"  Paper Size: {pageSetup.PaperSize}");
        sb.AppendLine($"  Fit To Pages Wide: {pageSetup.FitToPagesWide}");
        sb.AppendLine($"  Fit To Pages Tall: {pageSetup.FitToPagesTall}");

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    ///     Edits worksheet properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing various property values</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> EditSheetPropertiesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var name = ArgumentHelper.GetStringNullable(arguments, "name");
        var isVisible = ArgumentHelper.GetBoolNullable(arguments, "isVisible");
        var tabColor = ArgumentHelper.GetStringNullable(arguments, "tabColor");
        var isSelected = ArgumentHelper.GetBoolNullable(arguments, "isSelected");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (!string.IsNullOrEmpty(name)) worksheet.Name = name;

        if (isVisible.HasValue) worksheet.IsVisible = isVisible.Value;

        if (!string.IsNullOrWhiteSpace(tabColor))
        {
            var color = ColorHelper.ParseColor(tabColor);
            worksheet.TabColor = color;
        }

        if (isSelected.HasValue && isSelected.Value) workbook.Worksheets.ActiveSheetIndex = sheetIndex;

        workbook.Save(outputPath);
        return await Task.FromResult($"Sheet {sheetIndex} properties updated: {outputPath}");
    }

    /// <summary>
    ///     Gets information about all worksheets
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Formatted string with sheet information</returns>
    private async Task<string> GetSheetInfoAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");

        using var workbook = new Workbook(path);
        var result = new StringBuilder();

        result.AppendLine("=== Excel Workbook Information ===\n");
        result.AppendLine($"Total worksheets: {workbook.Worksheets.Count}\n");

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
                throw new ArgumentException(
                    $"Worksheet index {sheetIndex.Value} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

            var worksheet = workbook.Worksheets[sheetIndex.Value];
            AppendSheetInfo(result, worksheet, sheetIndex.Value);
        }
        else
        {
            for (var i = 0; i < workbook.Worksheets.Count; i++)
            {
                AppendSheetInfo(result, workbook.Worksheets[i], i);
                if (i < workbook.Worksheets.Count - 1) result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendSheetInfo(StringBuilder result, Worksheet worksheet, int index)
    {
        result.AppendLine($"[Worksheet {index}: {worksheet.Name}]");
        result.AppendLine($"  Visibility: {worksheet.VisibilityType}");
        result.AppendLine($"  Max row: {worksheet.Cells.MaxDataRow + 1}");
        result.AppendLine($"  Max column: {worksheet.Cells.MaxDataColumn + 1}");
        result.AppendLine($"  Used range: {worksheet.Cells.MaxRow + 1} rows × {worksheet.Cells.MaxColumn + 1} columns");
        result.AppendLine($"  Page orientation: {worksheet.PageSetup.Orientation}");
        result.AppendLine($"  Paper size: {worksheet.PageSetup.PaperSize}");
        result.AppendLine($"  Freeze panes: row {worksheet.FirstVisibleRow}, column {worksheet.FirstVisibleColumn}");
    }
}