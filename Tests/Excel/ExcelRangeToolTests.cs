using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelRangeToolTests : ExcelTestBase
{
    private readonly ExcelRangeTool _tool = new();

    [Fact]
    public async Task WriteRange_ShouldWriteDataToRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_write_range.xlsx");
        var outputPath = CreateTestFilePath("test_write_range_output.xlsx");
        var arguments = CreateArguments("write", workbookPath, outputPath);
        arguments["startCell"] = "A1";
        arguments["data"] = new JsonArray
        {
            new JsonArray { "A", "B" },
            new JsonArray { "C", "D" }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("A", worksheet.Cells["A1"].Value);
        Assert.Equal("B", worksheet.Cells["B1"].Value);
        Assert.Equal("C", worksheet.Cells["A2"].Value);
        Assert.Equal("D", worksheet.Cells["B2"].Value);
    }

    [Fact]
    public async Task GetRange_ShouldReturnRangeData()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_range.xlsx", 3);
        var arguments = CreateArguments("get", workbookPath);
        arguments["range"] = "A1:B2";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("R1C1", result);
        Assert.Contains("R1C2", result);
    }

    [Fact]
    public async Task ClearRange_ShouldClearRangeContent()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_clear_range.xlsx", 3);
        var outputPath = CreateTestFilePath("test_clear_range_output.xlsx");
        var arguments = CreateArguments("clear", workbookPath, outputPath);
        arguments["range"] = "A1:B2";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Cleared cells should be empty
        var a1Value = worksheet.Cells["A1"].Value;
        Assert.True(a1Value == null || a1Value.ToString() == "",
            $"Cell A1 should be cleared, got: {a1Value}");
    }

    [Fact]
    public async Task CopyRange_ShouldCopyRangeToDestination()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_copy_range.xlsx", 3);
        var outputPath = CreateTestFilePath("test_copy_range_output.xlsx");
        var arguments = CreateArguments("copy", workbookPath, outputPath);
        arguments["sourceRange"] = "A1:B2";
        arguments["destCell"] = "C1";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Verify data was copied
        var sourceA1 = worksheet.Cells["A1"].Value?.ToString() ?? "";
        var destC1 = worksheet.Cells["C1"].Value?.ToString() ?? "";
        Assert.Equal(sourceA1, destC1);
    }

    [Fact]
    public async Task MoveRange_ShouldMoveRangeToDestination()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_move_range.xlsx", 3);
        var sourceA1Value = new Workbook(workbookPath).Worksheets[0].Cells["A1"].Value;
        var outputPath = CreateTestFilePath("test_move_range_output.xlsx");
        var arguments = CreateArguments("move", workbookPath, outputPath);
        arguments["sourceRange"] = "A1:B2";
        arguments["destCell"] = "C1";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Verify data was moved (destination should have data)
        var destC1 = worksheet.Cells["C1"].Value;
        Assert.Equal(sourceA1Value, destC1);
        // Source should be cleared (moved, not copied)
        var sourceA1 = worksheet.Cells["A1"].Value;
        Assert.True(sourceA1 == null || sourceA1.ToString() == "",
            $"Source cell A1 should be cleared after move, got: {sourceA1}");
    }

    [Fact]
    public async Task EditRange_ShouldEditRangeData()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_range.xlsx", 3);
        var outputPath = CreateTestFilePath("test_edit_range_output.xlsx");
        var arguments = CreateArguments("edit", workbookPath, outputPath);
        arguments["range"] = "A1:B2";
        arguments["data"] = new JsonArray
        {
            new JsonArray { "X", "Y" },
            new JsonArray { "Z", "W" }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("X", worksheet.Cells["A1"].Value);
        Assert.Equal("Y", worksheet.Cells["B1"].Value);
        Assert.Equal("Z", worksheet.Cells["A2"].Value);
        Assert.Equal("W", worksheet.Cells["B2"].Value);
    }

    [Fact]
    public async Task CopyFormat_ShouldCopyFormatOnly()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_copy_format.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var sourceCell = worksheet.Cells["A1"];
        sourceCell.Value = "Test";
        var style = sourceCell.GetStyle();
        style.Font.IsBold = true;
        style.Font.Size = 14;
        sourceCell.SetStyle(style);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_format_output.xlsx");
        var arguments = CreateArguments("copy_format", workbookPath, outputPath);
        arguments["sourceRange"] = "A1";
        arguments["destCell"] = "B1";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var destStyle = resultWorksheet.Cells["B1"].GetStyle();
        Assert.True(destStyle.Font.IsBold, "Format should be copied (bold should be true)");
        Assert.Equal(14, destStyle.Font.Size);
    }
}