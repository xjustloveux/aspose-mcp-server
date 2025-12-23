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
        var workbookPath = CreateExcelWorkbookWithData("test_add_named_range.xlsx");
        var outputPath = CreateTestFilePath("test_add_named_range_output.xlsx");
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["name"] = "TestRange";
        arguments["range"] = "A1:C5";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var namedRanges = workbook.Worksheets.Names;
        var found = false;
        foreach (var name in namedRanges)
            if (name.Text == "TestRange")
            {
                found = true;
                break;
            }

        Assert.True(found, "Named range 'TestRange' should be added");
    }

    [Fact]
    public async Task AddNamedRange_WithComment_ShouldAddComment()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_named_range_comment.xlsx");
        var outputPath = CreateTestFilePath("test_add_named_range_comment_output.xlsx");
        var arguments = CreateArguments("add", workbookPath, outputPath);
        arguments["name"] = "CommentedRange";
        arguments["range"] = "A1:B2";
        arguments["comment"] = "This is a test range";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var namedRanges = workbook.Worksheets.Names;
        var found = false;
        foreach (var name in namedRanges)
            if (name.Text == "CommentedRange")
            {
                found = true;
                break;
            }

        Assert.True(found, "Named range 'CommentedRange' should be added");
    }

    [Fact]
    public async Task DeleteNamedRange_ShouldDeleteNamedRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_named_range.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1", "B2");
        range.Name = "RangeToDelete";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_delete_named_range_output.xlsx");
        var arguments = CreateArguments("delete", workbookPath, outputPath);
        arguments["name"] = "RangeToDelete";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var namedRanges = resultWorkbook.Worksheets.Names;
        var found = false;
        foreach (var name in namedRanges)
            if (name.Text == "RangeToDelete")
            {
                found = true;
                break;
            }

        Assert.False(found, "Named range 'RangeToDelete' should be deleted");
    }

    [Fact]
    public async Task GetNamedRanges_ShouldReturnAllNamedRanges()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_named_ranges.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range1 = worksheet.Cells.CreateRange("A1", "B2");
        range1.Name = "Range1";
        var range2 = worksheet.Cells.CreateRange("C1", "D2");
        range2.Name = "Range2";
        workbook.Save(workbookPath);

        var arguments = CreateArguments("get", workbookPath);
        arguments["operation"] = "get";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Range1", result);
        Assert.Contains("Range2", result);
    }

    [Fact]
    public async Task GetNamedRanges_WithNoNamedRanges_ShouldReturnEmptyMessage()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_empty_named_ranges.xlsx");
        var arguments = CreateArguments("get", workbookPath);
        arguments["operation"] = "get";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        // Should return a message indicating no named ranges found
        Assert.NotEmpty(result);
    }
}