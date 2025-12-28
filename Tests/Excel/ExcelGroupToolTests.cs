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
        var workbookPath = CreateExcelWorkbookWithData("test_group_rows.xlsx", 10, 5);
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Rows 1-3 grouped", result);
        Assert.Contains(outputPath, result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Cells.Rows[1].GroupLevel > 0 || worksheet.Cells.Rows[2].GroupLevel > 0);
    }

    [Fact]
    public async Task GroupRows_WithCollapsed_ShouldGroupAndCollapse()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_rows_collapsed.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_group_rows_collapsed_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "group_rows",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["startRow"] = 1,
            ["endRow"] = 3,
            ["isCollapsed"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Rows 1-3 grouped", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task GroupRows_MissingStartRow_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_missing_start.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "group_rows",
            ["path"] = workbookPath,
            ["endRow"] = 3
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("requires parameter 'startRow'", exception.Message);
    }

    [Fact]
    public async Task GroupRows_MissingEndRow_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_missing_end.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "group_rows",
            ["path"] = workbookPath,
            ["startRow"] = 1
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("requires parameter 'endRow'", exception.Message);
    }

    [Fact]
    public async Task GroupRows_StartGreaterThanEnd_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_invalid_range.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "group_rows",
            ["path"] = workbookPath,
            ["startRow"] = 5,
            ["endRow"] = 2
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("cannot be greater than", exception.Message);
    }

    [Fact]
    public async Task GroupRows_NegativeStart_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_negative.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "group_rows",
            ["path"] = workbookPath,
            ["startRow"] = -1,
            ["endRow"] = 3
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("cannot be negative", exception.Message);
    }

    [Fact]
    public async Task UngroupRows_ShouldUngroupRows()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_rows.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells.GroupRows(1, 3, false);
            workbook.Save(workbookPath);
        }

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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Rows 1-3 ungrouped", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task GroupColumns_ShouldGroupColumns()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_columns.xlsx", 5, 10);
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Columns 1-3 grouped", result);
        Assert.True(File.Exists(outputPath));

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Cells.Columns[1].GroupLevel > 0 || worksheet.Cells.Columns[2].GroupLevel > 0);
    }

    [Fact]
    public async Task GroupColumns_MissingStartColumn_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_col_missing.xlsx", 5, 10);
        var arguments = new JsonObject
        {
            ["operation"] = "group_columns",
            ["path"] = workbookPath,
            ["endColumn"] = 3
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("requires parameter 'startColumn'", exception.Message);
    }

    [Fact]
    public async Task GroupColumns_StartGreaterThanEnd_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_col_invalid.xlsx", 5, 10);
        var arguments = new JsonObject
        {
            ["operation"] = "group_columns",
            ["path"] = workbookPath,
            ["startColumn"] = 5,
            ["endColumn"] = 2
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("cannot be greater than", exception.Message);
    }

    [Fact]
    public async Task UngroupColumns_ShouldUngroupColumns()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_columns.xlsx", 5, 10);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells.GroupColumns(1, 3, false);
            workbook.Save(workbookPath);
        }

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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Columns 1-3 ungrouped", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_invalid_op.xlsx", 5, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "invalid",
            ["path"] = workbookPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task GroupRows_WithSheetIndex_ShouldGroupCorrectSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_sheet_index.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells[0, 0].Value = "Test";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_group_sheet_index_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "group_rows",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 1,
            ["startRow"] = 0,
            ["endRow"] = 2
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("sheet 1", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task GroupRows_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_invalid_sheet.xlsx", 10, 5);
        var arguments = new JsonObject
        {
            ["operation"] = "group_rows",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["startRow"] = 0,
            ["endRow"] = 2
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GroupRows_SingleRow_ShouldSucceed()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_group_single_row.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_group_single_row_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "group_rows",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["startRow"] = 2,
            ["endRow"] = 2
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Rows 2-2 grouped", result);
    }
}