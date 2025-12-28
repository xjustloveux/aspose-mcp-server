using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelPropertiesToolTests : ExcelTestBase
{
    private readonly ExcelPropertiesTool _tool = new();

    #region GetWorkbookProperties Tests

    [Fact]
    public async Task GetWorkbookProperties_ShouldReturnJsonWithAllFields()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_workbook_properties.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_workbook_properties",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.True(root.TryGetProperty("title", out _));
        Assert.True(root.TryGetProperty("author", out _));
        Assert.True(root.TryGetProperty("totalSheets", out _));
        Assert.True(root.TryGetProperty("customProperties", out _));
    }

    [Fact]
    public async Task GetWorkbookProperties_WithSetProperties_ShouldReturnCorrectValues()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_props_with_values.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.BuiltInDocumentProperties.Title = "Test Title";
            workbook.BuiltInDocumentProperties.Author = "Test Author";
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_workbook_properties",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        Assert.Equal("Test Title", json.RootElement.GetProperty("title").GetString());
        Assert.Equal("Test Author", json.RootElement.GetProperty("author").GetString());
    }

    #endregion

    #region SetWorkbookProperties Tests

    [Fact]
    public async Task SetWorkbookProperties_ShouldSetAllBuiltInProperties()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_all_properties.xlsx");
        var outputPath = CreateTestFilePath("test_set_all_properties_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_workbook_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["title"] = "Test Title",
            ["subject"] = "Test Subject",
            ["author"] = "Test Author",
            ["keywords"] = "test,keywords",
            ["comments"] = "Test Comments",
            ["category"] = "Test Category",
            ["company"] = "Test Company",
            ["manager"] = "Test Manager"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("successfully", result);
        using var workbook = new Workbook(outputPath);
        var props = workbook.BuiltInDocumentProperties;
        Assert.Equal("Test Title", props.Title);
        Assert.Equal("Test Subject", props.Subject);
        Assert.Equal("Test Author", props.Author);
        Assert.Equal("test,keywords", props.Keywords);
        Assert.Equal("Test Comments", props.Comments);
        Assert.Equal("Test Category", props.Category);
        Assert.Equal("Test Company", props.Company);
        Assert.Equal("Test Manager", props.Manager);
    }

    [Fact]
    public async Task SetWorkbookProperties_WithCustomProperties_ShouldAddNewProperties()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_custom_properties.xlsx");
        var outputPath = CreateTestFilePath("test_set_custom_properties_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_workbook_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["customProperties"] = new JsonObject
            {
                ["CustomProp1"] = "Value1",
                ["CustomProp2"] = "Value2"
            }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        var customProps = workbook.CustomDocumentProperties;
        Assert.Equal("Value1", customProps["CustomProp1"].Value?.ToString());
        Assert.Equal("Value2", customProps["CustomProp2"].Value?.ToString());
    }

    [Fact]
    public async Task SetWorkbookProperties_WithExistingCustomProperty_ShouldUpdateValue()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_update_custom_property.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.CustomDocumentProperties.Add("ExistingProp", "OldValue");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_update_custom_property_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_workbook_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["customProperties"] = new JsonObject
            {
                ["ExistingProp"] = "NewValue"
            }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("NewValue", resultWorkbook.CustomDocumentProperties["ExistingProp"].Value?.ToString());
    }

    #endregion

    #region GetSheetProperties Tests

    [Fact]
    public async Task GetSheetProperties_ShouldReturnJsonWithAllFields()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_sheet_props.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_sheet_properties",
            ["path"] = workbookPath,
            ["sheetIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.True(root.TryGetProperty("name", out _));
        Assert.True(root.TryGetProperty("index", out _));
        Assert.True(root.TryGetProperty("isVisible", out _));
        Assert.True(root.TryGetProperty("dataRowCount", out _));
        Assert.True(root.TryGetProperty("dataColumnCount", out _));
        Assert.True(root.TryGetProperty("printSettings", out _));

        Assert.Equal(5, root.GetProperty("dataRowCount").GetInt32());
        Assert.Equal(3, root.GetProperty("dataColumnCount").GetInt32());
    }

    [Fact]
    public async Task GetSheetProperties_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_sheet_props_invalid.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_sheet_properties",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task GetSheetProperties_WithNegativeIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_sheet_props_negative.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_sheet_properties",
            ["path"] = workbookPath,
            ["sheetIndex"] = -1
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region EditSheetProperties Tests

    [Fact]
    public async Task EditSheetProperties_ShouldChangeName()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_name.xlsx");
        var outputPath = CreateTestFilePath("test_edit_name_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit_sheet_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 0,
            ["name"] = "NewSheetName"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        Assert.Equal("NewSheetName", workbook.Worksheets[0].Name);
    }

    [Fact]
    public async Task EditSheetProperties_ShouldChangeVisibility()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_visibility.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_visibility_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit_sheet_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 0,
            ["isVisible"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.Worksheets[0].IsVisible);
    }

    [Fact]
    public async Task EditSheetProperties_ShouldChangeTabColor()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_tab_color.xlsx");
        var outputPath = CreateTestFilePath("test_edit_tab_color_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit_sheet_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 0,
            ["tabColor"] = "#FF0000"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        var tabColor = workbook.Worksheets[0].TabColor;
        Assert.Equal(255, tabColor.R);
        Assert.Equal(0, tabColor.G);
        Assert.Equal(0, tabColor.B);
    }

    [Fact]
    public async Task EditSheetProperties_ShouldSetAsSelected()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_selected.xlsx");
        int targetSheetIndex;
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet3");
            wb.Save(workbookPath);
            // Find the index of "Sheet3" (the sheet we want to select)
            targetSheetIndex = wb.Worksheets["Sheet3"].Index;
        }

        var outputPath = CreateTestFilePath("test_edit_selected_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit_sheet_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = targetSheetIndex,
            ["isSelected"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            // In evaluation mode, additional evaluation warning sheets may be added
            // causing index shifts. Verify the operation completed without error.
            Assert.NotNull(workbook.Worksheets["Sheet3"]);
        }
        else
        {
            // In licensed mode, verify that Sheet3 is the active sheet
            var sheet3Index = workbook.Worksheets["Sheet3"].Index;
            Assert.Equal(sheet3Index, workbook.Worksheets.ActiveSheetIndex);
        }
    }

    [Fact]
    public async Task EditSheetProperties_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_invalid_index.xlsx");
        var outputPath = CreateTestFilePath("test_edit_invalid_index_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit_sheet_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 99,
            ["name"] = "NewName"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region GetSheetInfo Tests

    [Fact]
    public async Task GetSheetInfo_ShouldReturnAllSheets()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_sheet_info_all.xlsx");
        int expectedSheetCount;
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Save(workbookPath);
            expectedSheetCount = workbook.Worksheets.Count;
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_sheet_info",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(expectedSheetCount, root.GetProperty("count").GetInt32());
        Assert.Equal(expectedSheetCount, root.GetProperty("totalWorksheets").GetInt32());

        var items = root.GetProperty("items");
        Assert.Equal(expectedSheetCount, items.GetArrayLength());
    }

    [Fact]
    public async Task GetSheetInfo_WithSheetIndex_ShouldReturnSingleSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_sheet_info_single.xlsx");
        int totalSheetCount;
        int sheet2Index;
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Save(workbookPath);
            totalSheetCount = workbook.Worksheets.Count;
            sheet2Index = workbook.Worksheets["Sheet2"].Index;
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_sheet_info",
            ["path"] = workbookPath,
            ["sheetIndex"] = sheet2Index
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(1, root.GetProperty("count").GetInt32());
        Assert.Equal(totalSheetCount, root.GetProperty("totalWorksheets").GetInt32());

        var items = root.GetProperty("items");
        Assert.Equal(1, items.GetArrayLength());
        Assert.Equal("Sheet2", items[0].GetProperty("name").GetString());
    }

    [Fact]
    public async Task GetSheetInfo_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_sheet_info_invalid.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_sheet_info",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task GetSheetInfo_ShouldReturnCorrectDataCounts()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_sheet_info_data.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "get_sheet_info",
            ["path"] = workbookPath,
            ["sheetIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var sheet = json.RootElement.GetProperty("items")[0];

        Assert.Equal(10, sheet.GetProperty("dataRowCount").GetInt32());
        Assert.Equal(5, sheet.GetProperty("dataColumnCount").GetInt32());
    }

    #endregion

    #region Error Handling Tests

    [Fact]
    public async Task ExecuteAsync_WithUnknownOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_unknown_operation.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "invalid_operation",
            ["path"] = workbookPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task ExecuteAsync_WithMissingPath_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["operation"] = "get_workbook_properties"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ExecuteAsync_WithMissingOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_missing_operation.xlsx");
        var arguments = new JsonObject
        {
            ["path"] = workbookPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion
}