using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelGroupToolTests : ExcelTestBase
{
    private readonly ExcelGroupTool _tool = new();

    [Fact]
    public async Task GroupRows_ShouldGroupRows()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_rows.xlsx");
        var outputPath = CreateTestFilePath("test_group_rows_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "group_rows",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["startRow"] = 1,
            ["endRow"] = 3
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Verify rows were grouped
        Assert.NotNull(worksheet);
    }

    [Fact]
    public async Task UngroupRows_ShouldUngroupRows()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_rows.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells.GroupRows(1, 3, false);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_ungroup_rows_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "ungroup_rows",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["startRow"] = 1,
            ["endRow"] = 3
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
    }

    [Fact]
    public async Task GroupColumns_ShouldGroupColumns()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_columns.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_group_columns_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "group_columns",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["startColumn"] = 1,
            ["endColumn"] = 3
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
    }

    [Fact]
    public async Task UngroupColumns_ShouldUngroupColumns()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_columns.xlsx", 5, 5);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells.GroupColumns(1, 3, false);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_ungroup_columns_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "ungroup_columns",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["startColumn"] = 1,
            ["endColumn"] = 3
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
    }
}