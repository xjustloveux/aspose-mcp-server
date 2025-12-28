using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelMergeCellsToolTests : ExcelTestBase
{
    private readonly ExcelMergeCellsTool _tool = new();

    [Fact]
    public async Task MergeCells_ShouldMergeRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_merge_cells.xlsx", 3);
        var outputPath = CreateTestFilePath("test_merge_cells_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:C1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Range A1:C1 merged", result);
        Assert.Contains("1 rows x 3 columns", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Cells.MergedCells.Count > 0);
    }

    [Fact]
    public async Task MergeCells_MultipleRows_ShouldMerge()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_merge_multi.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_merge_multi_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:B3"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Range A1:B3 merged", result);
        Assert.Contains("3 rows x 2 columns", result);

        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public async Task MergeCells_SingleCell_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_merge_single.xlsx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["path"] = workbookPath,
            ["range"] = "A1"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Cannot merge a single cell", exception.Message);
    }

    [Fact]
    public async Task MergeCells_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_merge_invalid_sheet.xlsx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["range"] = "A1:C1"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task UnmergeCells_ShouldUnmergeRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_unmerge_cells.xlsx", 3);
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3);
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_unmerge_cells_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unmerge",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:C1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Range A1:C1 unmerged", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Empty(resultWorkbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public async Task UnmergeCells_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_unmerge_invalid_sheet.xlsx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "unmerge",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["range"] = "A1:C1"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetMergedCells_ShouldReturnMergedCells()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_merged_cells.xlsx", 3);
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells[0, 0].Value = "Header";
            workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3);
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(1, root.GetProperty("count").GetInt32());
        var items = root.GetProperty("items");
        Assert.Equal(1, items.GetArrayLength());

        var firstItem = items[0];
        Assert.Equal("A1:C1", firstItem.GetProperty("range").GetString());
        Assert.Equal("A1", firstItem.GetProperty("startCell").GetString());
        Assert.Equal("C1", firstItem.GetProperty("endCell").GetString());
        Assert.Equal(1, firstItem.GetProperty("rowCount").GetInt32());
        Assert.Equal(3, firstItem.GetProperty("columnCount").GetInt32());
        Assert.Equal("Header", firstItem.GetProperty("value").GetString());
    }

    [Fact]
    public async Task GetMergedCells_EmptyWorksheet_ShouldReturnEmptyResult()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_no_merged.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal("No merged cells found", root.GetProperty("message").GetString());
    }

    [Fact]
    public async Task GetMergedCells_MultipleMergedRanges_ShouldReturnAll()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_multi_merged.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3); // A1:C1
            workbook.Worksheets[0].Cells.Merge(2, 0, 2, 2); // A3:B4
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(2, root.GetProperty("count").GetInt32());
        Assert.Equal(2, root.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public async Task GetMergedCells_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_op.xlsx");
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
    public async Task MergeCells_WithSheetIndex_ShouldMergeCorrectSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_merge_sheet_index.xlsx", 3);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells[0, 0].Value = "Test";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_merge_sheet_index_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 1,
            ["range"] = "A1:C1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("merged", result);

        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Cells.MergedCells);
        Assert.Single(workbook.Worksheets[1].Cells.MergedCells);
    }
}