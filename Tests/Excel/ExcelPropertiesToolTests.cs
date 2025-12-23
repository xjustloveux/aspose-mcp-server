using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelPropertiesToolTests : ExcelTestBase
{
    private readonly ExcelPropertiesTool _tool = new();

    [Fact]
    public async Task GetWorkbookProperties_ShouldReturnProperties()
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
        Assert.NotEmpty(result);
        Assert.Contains("Properties", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SetWorkbookProperties_ShouldSetProperties()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_workbook_properties.xlsx");
        var outputPath = CreateTestFilePath("test_set_workbook_properties_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_workbook_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["title"] = "Test Workbook",
            ["author"] = "Test Author"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        Assert.Equal("Test Workbook", workbook.BuiltInDocumentProperties["Title"].ToString());
        Assert.Equal("Test Author", workbook.BuiltInDocumentProperties["Author"].ToString());
    }

    [Fact]
    public async Task GetSheetProperties_ShouldReturnSheetProperties()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_sheet_properties.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_sheet_properties",
            ["path"] = workbookPath,
            ["sheetIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public async Task GetSheetInfo_ShouldReturnSheetInfo()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_sheet_info.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_sheet_info",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Sheet", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task EditSheetProperties_ShouldEditSheetProperties()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_sheet_properties.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets.Add("OriginalName");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_sheet_properties_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit_sheet_properties",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 1,
            ["name"] = "EditedName"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var sheet = resultWorkbook.Worksheets["EditedName"];
        Assert.NotNull(sheet);
    }
}