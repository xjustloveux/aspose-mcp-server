using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelNamedRangeToolTests : ExcelTestBase
{
    private readonly ExcelNamedRangeTool _tool = new();

    [Fact]
    public async Task AddNamedRange_ShouldAddNamedRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_add_named_range.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_named_range_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["name"] = "TestRange",
            ["range"] = "A1:C5"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Named range 'TestRange' added", result);
        Assert.Contains("reference:", result);

        using var workbook = new Workbook(outputPath);
        Assert.NotNull(workbook.Worksheets.Names["TestRange"]);
    }

    [Fact]
    public async Task AddNamedRange_WithComment_ShouldAddComment()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_named_range_comment.xlsx");
        var outputPath = CreateTestFilePath("test_add_named_range_comment_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["name"] = "CommentedRange",
            ["range"] = "A1:B2",
            ["comment"] = "This is a test range"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Named range 'CommentedRange' added", result);

        using var workbook = new Workbook(outputPath);
        var namedRange = workbook.Worksheets.Names["CommentedRange"];
        Assert.NotNull(namedRange);
        Assert.Equal("This is a test range", namedRange.Comment);
    }

    [Fact]
    public async Task AddNamedRange_SingleCell_ShouldAddRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_single_cell.xlsx");
        var outputPath = CreateTestFilePath("test_add_single_cell_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["name"] = "SingleCell",
            ["range"] = "A1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Named range 'SingleCell' added", result);

        using var workbook = new Workbook(outputPath);
        Assert.NotNull(workbook.Worksheets.Names["SingleCell"]);
    }

    [Fact]
    public async Task AddNamedRange_WithSheetReference_ShouldAddToCorrectSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_sheet_ref.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("DataSheet");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_sheet_ref_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["name"] = "SheetRange",
            ["range"] = "DataSheet!A1:C5"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Named range 'SheetRange' added", result);
        Assert.Contains("DataSheet", result);

        using var workbook = new Workbook(outputPath);
        var namedRange = workbook.Worksheets.Names["SheetRange"];
        Assert.NotNull(namedRange);
        Assert.Contains("DataSheet", namedRange.RefersTo);
    }

    [Fact]
    public async Task AddNamedRange_DuplicateName_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_duplicate.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var range = wb.Worksheets[0].Cells.CreateRange("A1", "B2");
            range.Name = "ExistingRange";
            wb.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["name"] = "ExistingRange",
            ["range"] = "C1:D2"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("already exists", exception.Message);
    }

    [Fact]
    public async Task AddNamedRange_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["name"] = "InvalidSheet",
            ["range"] = "A1:B2",
            ["sheetIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task AddNamedRange_InvalidSheetReference_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet_ref.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["name"] = "InvalidRef",
            ["range"] = "NonExistentSheet!A1:B2"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public async Task DeleteNamedRange_ShouldDeleteNamedRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_named_range.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var range = workbook.Worksheets[0].Cells.CreateRange("A1", "B2");
            range.Name = "RangeToDelete";
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_delete_named_range_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["name"] = "RangeToDelete"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Named range 'RangeToDelete' deleted", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Null(resultWorkbook.Worksheets.Names["RangeToDelete"]);
    }

    [Fact]
    public async Task DeleteNamedRange_NonExistent_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_nonexistent.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["name"] = "NonExistentRange"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("does not exist", exception.Message);
    }

    [Fact]
    public async Task GetNamedRanges_ShouldReturnAllNamedRanges()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_named_ranges.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var range1 = workbook.Worksheets[0].Cells.CreateRange("A1", "B2");
            range1.Name = "Range1";
            var range2 = workbook.Worksheets[0].Cells.CreateRange("C1", "D2");
            range2.Name = "Range2";
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
        var items = root.GetProperty("items");
        Assert.Equal(2, items.GetArrayLength());

        Assert.Contains("Range1", result);
        Assert.Contains("Range2", result);
    }

    [Fact]
    public async Task GetNamedRanges_WithNoNamedRanges_ShouldReturnEmptyMessage()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_empty_named_ranges.xlsx");
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
        Assert.Equal("No named ranges found", root.GetProperty("message").GetString());
    }

    [Fact]
    public async Task GetNamedRanges_ShouldIncludeAllProperties()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_properties.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var range = workbook.Worksheets[0].Cells.CreateRange("A1", "B2");
            range.Name = "DetailedRange";
            var namedRange = workbook.Worksheets.Names["DetailedRange"];
            namedRange.Comment = "Test comment";
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
        var items = json.RootElement.GetProperty("items");
        var firstItem = items[0];

        Assert.True(firstItem.TryGetProperty("name", out _));
        Assert.True(firstItem.TryGetProperty("reference", out _));
        Assert.True(firstItem.TryGetProperty("comment", out _));
        Assert.True(firstItem.TryGetProperty("isVisible", out _));
        Assert.Equal("DetailedRange", firstItem.GetProperty("name").GetString());
        Assert.Equal("Test comment", firstItem.GetProperty("comment").GetString());
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
    public async Task AddNamedRange_WithSheetIndex_ShouldAddToCorrectSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_with_sheet_index.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_with_sheet_index_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["name"] = "Sheet2Range",
            ["range"] = "A1:C5",
            ["sheetIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Named range 'Sheet2Range' added", result);

        using var workbook = new Workbook(outputPath);
        var namedRange = workbook.Worksheets.Names["Sheet2Range"];
        Assert.NotNull(namedRange);
        Assert.Contains("Sheet2", namedRange.RefersTo);
    }
}