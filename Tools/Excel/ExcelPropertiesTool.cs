using System.Text.Json;
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
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "get_workbook_properties" => await GetWorkbookPropertiesAsync(path),
            "set_workbook_properties" => await SetWorkbookPropertiesAsync(path, outputPath, arguments),
            "get_sheet_properties" => await GetSheetPropertiesAsync(path, sheetIndex),
            "edit_sheet_properties" => await EditSheetPropertiesAsync(path, outputPath, sheetIndex, arguments),
            "get_sheet_info" => await GetSheetInfoAsync(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets workbook properties
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <returns>JSON string with workbook properties</returns>
    private Task<string> GetWorkbookPropertiesAsync(string path)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var props = workbook.BuiltInDocumentProperties;
            var customProps = workbook.CustomDocumentProperties;

            var customPropsList = new List<object>();
            if (customProps.Count > 0)
                foreach (var prop in customProps)
                    customPropsList.Add(new { name = prop.Name, value = prop.Value?.ToString() });

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
        });
    }

    /// <summary>
    ///     Sets workbook properties
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing various property values</param>
    /// <returns>Success message</returns>
    private Task<string> SetWorkbookPropertiesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
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
            return $"Workbook properties updated successfully. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets worksheet properties
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>JSON string with worksheet properties</returns>
    private Task<string> GetSheetPropertiesAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pageSetup = worksheet.PageSetup;

            var result = new
            {
                name = worksheet.Name,
                index = sheetIndex,
                isVisible = worksheet.IsVisible,
                tabColor = worksheet.TabColor.ToString(),
                isSelected = workbook.Worksheets.ActiveSheetIndex == sheetIndex,
                maxDataRow = worksheet.Cells.MaxDataRow,
                maxDataColumn = worksheet.Cells.MaxDataColumn,
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
        });
    }

    /// <summary>
    ///     Edits worksheet properties
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing various property values</param>
    /// <returns>Success message</returns>
    private Task<string> EditSheetPropertiesAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
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
            return $"Sheet {sheetIndex} properties updated successfully. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets information about all worksheets
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>JSON string with sheet information</returns>
    private Task<string> GetSheetInfoAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");

            using var workbook = new Workbook(path);

            var sheetList = new List<object>();

            if (sheetIndex.HasValue)
            {
                if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
                    throw new ArgumentException(
                        $"Worksheet index {sheetIndex.Value} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

                var worksheet = workbook.Worksheets[sheetIndex.Value];
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
        });
    }

    private object CreateSheetInfo(Worksheet worksheet, int index)
    {
        return new
        {
            index,
            name = worksheet.Name,
            visibility = worksheet.VisibilityType.ToString(),
            maxRow = worksheet.Cells.MaxDataRow + 1,
            maxColumn = worksheet.Cells.MaxDataColumn + 1,
            usedRange = new
            {
                rows = worksheet.Cells.MaxRow + 1,
                columns = worksheet.Cells.MaxColumn + 1
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