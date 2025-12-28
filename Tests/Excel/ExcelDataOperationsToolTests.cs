using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelDataOperationsToolTests : ExcelTestBase
{
    private readonly ExcelDataOperationsTool _tool = new();

    [Fact]
    public async Task SortData_ShouldSortRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_sort.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "C";
        worksheet.Cells["A2"].Value = "A";
        worksheet.Cells["A3"].Value = "B";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_sort_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "sort",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A3",
            ["sortColumn"] = 0,
            ["ascending"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("A", resultWorksheet.Cells["A1"].Value);
        Assert.Equal("B", resultWorksheet.Cells["A2"].Value);
        Assert.Equal("C", resultWorksheet.Cells["A3"].Value);
    }

    [Fact]
    public async Task FindReplace_ShouldReplaceText()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_find_replace.xlsx", 3);
        var outputPath = CreateTestFilePath("test_find_replace_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "find_replace",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["findText"] = "R1C1",
            ["replaceText"] = "Replaced"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("Replaced", worksheet.Cells["A1"].Value);
    }

    [Fact]
    public async Task BatchWrite_ShouldWriteMultipleValues()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_batch_write.xlsx");
        var outputPath = CreateTestFilePath("test_batch_write_output.xlsx");
        var data = new JsonObject
        {
            ["A1"] = "Value1",
            ["B1"] = "Value2",
            ["A2"] = "Value3"
        };
        var arguments = new JsonObject
        {
            ["operation"] = "batch_write",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["data"] = data
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("Value1", worksheet.Cells["A1"].Value);
        Assert.Equal("Value2", worksheet.Cells["B1"].Value);
        Assert.Equal("Value3", worksheet.Cells["A2"].Value);
    }

    [Fact]
    public async Task GetContent_ShouldReturnContent()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_content.xlsx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "get_content",
            ["path"] = workbookPath,
            ["range"] = "A1:B2"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("R1C1", result);
    }

    [Fact]
    public async Task GetStatistics_ShouldReturnStatistics()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_statistics.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = 10;
        worksheet.Cells["A2"].Value = 20;
        worksheet.Cells["A3"].Value = 30;
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_statistics",
            ["path"] = workbookPath,
            ["range"] = "A1:A3"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Sum", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetUsedRange_ShouldReturnUsedRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_used_range.xlsx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "get_used_range",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Range", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SortData_WithHasHeader_ShouldSkipHeaderRow()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_sort_with_header.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["A2"].Value = "C";
        worksheet.Cells["A3"].Value = "A";
        worksheet.Cells["A4"].Value = "B";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_sort_with_header_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "sort",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A4",
            ["sortColumn"] = 0,
            ["ascending"] = true,
            ["hasHeader"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("Name", resultWorksheet.Cells["A1"].Value); // Header should remain
        Assert.Equal("A", resultWorksheet.Cells["A2"].Value);
        Assert.Equal("B", resultWorksheet.Cells["A3"].Value);
        Assert.Equal("C", resultWorksheet.Cells["A4"].Value);
    }

    [Fact]
    public async Task FindReplace_WithSubstring_ShouldNotLoopInfinitely()
    {
        // Arrange - Tests the fix for infinite loop when replaceText contains findText
        var workbookPath = CreateExcelWorkbook("test_find_replace_substring.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Apple";
        worksheet.Cells["A2"].Value = "Apple Pie";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_find_replace_substring_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "find_replace",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["findText"] = "Apple",
            ["replaceText"] = "AppleTree"
        };

        // Act - Should complete without infinite loop
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("2 replacements", result);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("AppleTree", resultWorksheet.Cells["A1"].Value);
        Assert.Equal("AppleTree Pie", resultWorksheet.Cells["A2"].Value);
    }

    [Fact]
    public async Task BatchWrite_WithArrayFormat_ShouldWriteValues()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_batch_write_array.xlsx");
        var outputPath = CreateTestFilePath("test_batch_write_array_output.xlsx");
        var data = new JsonArray
        {
            new JsonObject { ["cell"] = "A1", ["value"] = "Value1" },
            new JsonObject { ["cell"] = "B1", ["value"] = "Value2" }
        };
        var arguments = new JsonObject
        {
            ["operation"] = "batch_write",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["data"] = data
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("Value1", worksheet.Cells["A1"].Value);
        Assert.Equal("Value2", worksheet.Cells["B1"].Value);
    }

    [Fact]
    public async Task SortData_Descending_ShouldSortDescending()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_sort_desc.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "A";
        worksheet.Cells["A2"].Value = "C";
        worksheet.Cells["A3"].Value = "B";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_sort_desc_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "sort",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A3",
            ["sortColumn"] = 0,
            ["ascending"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("C", resultWorksheet.Cells["A1"].Value);
        Assert.Equal("B", resultWorksheet.Cells["A2"].Value);
        Assert.Equal("A", resultWorksheet.Cells["A3"].Value);
    }
}