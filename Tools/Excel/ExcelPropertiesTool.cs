using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Properties;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel properties (workbook properties, sheet properties, sheet info).
///     Merges: ExcelGetWorkbookPropertiesTool, ExcelSetWorkbookPropertiesTool, ExcelGetSheetPropertiesTool,
///     ExcelEditSheetPropertiesTool, ExcelGetSheetInfoTool.
/// </summary>
public class ExcelPropertiesTool : IAsposeTool
{
    private const string OperationGetWorkbookProperties = "get_workbook_properties";
    private const string OperationSetWorkbookProperties = "set_workbook_properties";
    private const string OperationGetSheetProperties = "get_sheet_properties";
    private const string OperationEditSheetProperties = "edit_sheet_properties";
    private const string OperationGetSheetInfo = "get_sheet_info";

    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description =>
        $@"Manage Excel properties. Supports 5 operations: {OperationGetWorkbookProperties}, {OperationSetWorkbookProperties}, {OperationGetSheetProperties}, {OperationEditSheetProperties}, {OperationGetSheetInfo}.

Usage examples:
- Get workbook properties: excel_properties(operation='{OperationGetWorkbookProperties}', path='book.xlsx')
- Set workbook properties: excel_properties(operation='{OperationSetWorkbookProperties}', path='book.xlsx', title='Title', author='Author')
- Get sheet properties: excel_properties(operation='{OperationGetSheetProperties}', path='book.xlsx', sheetIndex=0)
- Edit sheet properties: excel_properties(operation='{OperationEditSheetProperties}', path='book.xlsx', sheetIndex=0, name='New Name')
- Get sheet info: excel_properties(operation='{OperationGetSheetInfo}', path='book.xlsx')";

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
                description = $@"Operation to perform.
- '{OperationGetWorkbookProperties}': Get workbook properties (required params: path)
- '{OperationSetWorkbookProperties}': Set workbook properties (required params: path)
- '{OperationGetSheetProperties}': Get sheet properties (required params: path, sheetIndex)
- '{OperationEditSheetProperties}': Edit sheet properties (required params: path, sheetIndex)
- '{OperationGetSheetInfo}': Get sheet info (required params: path)",
                @enum = new[]
                {
                    OperationGetWorkbookProperties, OperationSetWorkbookProperties,
                    OperationGetSheetProperties, OperationEditSheetProperties, OperationGetSheetInfo
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
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            OperationGetWorkbookProperties => await GetWorkbookPropertiesAsync(path),
            OperationSetWorkbookProperties => await SetWorkbookPropertiesAsync(path, outputPath, arguments),
            OperationGetSheetProperties => await GetSheetPropertiesAsync(path, sheetIndex),
            OperationEditSheetProperties => await EditSheetPropertiesAsync(path, outputPath, sheetIndex, arguments),
            OperationGetSheetInfo => await GetSheetInfoAsync(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets workbook properties.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <returns>JSON string with workbook properties.</returns>
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
        });
    }

    /// <summary>
    ///     Sets workbook properties.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing various property values.</param>
    /// <returns>Success message.</returns>
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
                {
                    var value = kvp.Value?.GetValue<string>() ?? "";
                    var existingProp = FindCustomProperty(workbook.CustomDocumentProperties, kvp.Key);
                    if (existingProp != null)
                        workbook.CustomDocumentProperties.Remove(kvp.Key);
                    workbook.CustomDocumentProperties.Add(kvp.Key, value);
                }

            workbook.Save(outputPath);
            return $"Workbook properties updated successfully. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Finds a custom property by name.
    /// </summary>
    /// <param name="customProperties">Custom document properties collection.</param>
    /// <param name="name">Property name to find.</param>
    /// <returns>The property if found, otherwise null.</returns>
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
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>JSON string with worksheet properties.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range.</exception>
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
        });
    }

    /// <summary>
    ///     Edits worksheet properties.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing various property values.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range.</exception>
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
    ///     Gets information about all worksheets or a specific worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="arguments">JSON arguments (optional sheetIndex parameter).</param>
    /// <returns>JSON string with sheet information.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range.</exception>
    private Task<string> GetSheetInfoAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");

            using var workbook = new Workbook(path);

            var sheetList = new List<object>();

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
        });
    }

    /// <summary>
    ///     Creates sheet information object for JSON serialization.
    /// </summary>
    /// <param name="worksheet">Worksheet to get info from.</param>
    /// <param name="index">Worksheet index (0-based).</param>
    /// <returns>Anonymous object containing sheet information.</returns>
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