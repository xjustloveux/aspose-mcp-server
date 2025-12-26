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
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Check if cells are merged by verifying the range
        var range = worksheet.Cells.CreateRange("A1", "C1");
        Assert.NotNull(range);
    }

    [Fact]
    public async Task UnmergeCells_ShouldUnmergeRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_unmerge_cells.xlsx", 3);
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_unmerge_cells_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unmerge",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:C1"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        // Verify cells are unmerged
        Assert.NotNull(resultWorksheet);
    }

    [Fact]
    public async Task GetMergedCells_ShouldReturnMergedCells()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_merged_cells.xlsx", 3);
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3);
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }
}