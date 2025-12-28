using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelRangeToolTests : ExcelTestBase
{
    private readonly ExcelRangeTool _tool = new();

    #region Move Tests

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

    #endregion

    #region Write Tests

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
    public async Task WriteRange_WithObjectFormat_ShouldWriteData()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_write_object_format.xlsx");
        var outputPath = CreateTestFilePath("test_write_object_format_output.xlsx");
        var arguments = CreateArguments("write", workbookPath, outputPath);
        arguments["startCell"] = "A1";
        arguments["data"] = new JsonArray
        {
            new JsonObject { ["cell"] = "A1", ["value"] = "10" },
            new JsonObject { ["cell"] = "B2", ["value"] = "20" }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(10, Convert.ToDouble(worksheet.Cells["A1"].Value));
        Assert.Equal(20, Convert.ToDouble(worksheet.Cells["B2"].Value));
    }

    [Fact]
    public async Task WriteRange_WithNumericValues_ShouldStoreAsNumbers()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_write_numeric.xlsx");
        var outputPath = CreateTestFilePath("test_write_numeric_output.xlsx");
        var arguments = CreateArguments("write", workbookPath, outputPath);
        arguments["startCell"] = "A1";
        arguments["data"] = new JsonArray
        {
            new JsonArray { "100", "200.5", "true" }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(100, Convert.ToDouble(worksheet.Cells["A1"].Value));
        Assert.Equal(200.5, Convert.ToDouble(worksheet.Cells["B1"].Value));
        Assert.Equal(true, worksheet.Cells["C1"].Value);
    }

    [Fact]
    public async Task WriteRange_WithInvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_write_invalid_sheet.xlsx");
        var outputPath = CreateTestFilePath("test_write_invalid_sheet_output.xlsx");
        var arguments = CreateArguments("write", workbookPath, outputPath);
        arguments["sheetIndex"] = 99;
        arguments["startCell"] = "A1";
        arguments["data"] = new JsonArray { new JsonArray { "A" } };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region Get Tests

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
    public async Task GetRange_WithCalculateFormulas_ShouldRecalculate()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_calculate.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = 10;
            wb.Worksheets[0].Cells["A2"].Value = 20;
            wb.Worksheets[0].Cells["A3"].Formula = "=A1+A2";
            wb.Save(workbookPath);
        }

        var arguments = CreateArguments("get", workbookPath);
        arguments["range"] = "A3";
        arguments["calculateFormulas"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("30", result);
    }

    [Fact]
    public async Task GetRange_WithIncludeFormat_ShouldReturnFormatInfo()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_format.xlsx", 1);
        var arguments = CreateArguments("get", workbookPath);
        arguments["range"] = "A1";
        arguments["includeFormat"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("format", out var format));
        Assert.True(format.TryGetProperty("fontName", out _));
    }

    [Fact]
    public async Task GetRange_WithInvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_invalid_sheet.xlsx", 3);
        var arguments = CreateArguments("get", workbookPath);
        arguments["range"] = "A1:B2";
        arguments["sheetIndex"] = 99;

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region Clear Tests

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
    public async Task ClearRange_WithClearFormat_ShouldClearFormat()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_clear_format.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var boldStyle = wb.CreateStyle();
            boldStyle.Font.IsBold = true;
            wb.Worksheets[0].Cells["A1"].SetStyle(boldStyle);
            wb.Worksheets[0].Cells["A1"].Value = "Test";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_clear_format_output.xlsx");
        var arguments = CreateArguments("clear", workbookPath, outputPath);
        arguments["range"] = "A1";
        arguments["clearContent"] = false;
        arguments["clearFormat"] = true;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        var resultStyle = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.False(resultStyle.Font.IsBold);
    }

    #endregion

    #region Copy Tests

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
    public async Task CopyRange_WithValuesOnly_ShouldCopyValuesOnly()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_copy_values.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var style = wb.CreateStyle();
            style.Font.IsBold = true;
            wb.Worksheets[0].Cells["A1"].SetStyle(style);
            wb.Worksheets[0].Cells["A1"].Value = "Test";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_copy_values_output.xlsx");
        var arguments = CreateArguments("copy", workbookPath, outputPath);
        arguments["sourceRange"] = "A1";
        arguments["destCell"] = "B1";
        arguments["copyOptions"] = "Values";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Test", workbook.Worksheets[0].Cells["B1"].Value);
        var destStyle = workbook.Worksheets[0].Cells["B1"].GetStyle();
        Assert.False(destStyle.Font.IsBold);
    }

    #endregion

    #region Edit Tests

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
    public async Task EditRange_WithClearRange_ShouldClearBeforeEdit()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_clear.xlsx", 3);
        var outputPath = CreateTestFilePath("test_edit_clear_output.xlsx");
        var arguments = CreateArguments("edit", workbookPath, outputPath);
        arguments["range"] = "A1:C3";
        arguments["clearRange"] = true;
        arguments["data"] = new JsonArray
        {
            new JsonArray { "X" }
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("X", worksheet.Cells["A1"].Value);
        var b1 = worksheet.Cells["B1"].Value;
        Assert.True(b1 == null || b1.ToString() == "");
    }

    #endregion

    #region CopyFormat Tests

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

    [Fact]
    public async Task CopyFormat_WithCopyValue_ShouldCopyFormatAndValues()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_copy_format_value.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var style = wb.CreateStyle();
            style.Font.IsBold = true;
            wb.Worksheets[0].Cells["A1"].SetStyle(style);
            wb.Worksheets[0].Cells["A1"].Value = "Original";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_copy_format_value_output.xlsx");
        var arguments = CreateArguments("copy_format", workbookPath, outputPath);
        arguments["range"] = "A1";
        arguments["destCell"] = "B1";
        arguments["copyValue"] = true;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Original", workbook.Worksheets[0].Cells["B1"].Value);
        var destStyle = workbook.Worksheets[0].Cells["B1"].GetStyle();
        Assert.True(destStyle.Font.IsBold);
    }

    [Fact]
    public async Task CopyFormat_WithMissingDestination_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_copy_format_no_dest.xlsx");
        var outputPath = CreateTestFilePath("test_copy_format_no_dest_output.xlsx");
        var arguments = CreateArguments("copy_format", workbookPath, outputPath);
        arguments["range"] = "A1";

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("destRange or destCell is required", exception.Message);
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
            ["operation"] = "get"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task WriteRange_WithMissingStartCell_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_write_no_start.xlsx");
        var outputPath = CreateTestFilePath("test_write_no_start_output.xlsx");
        var arguments = CreateArguments("write", workbookPath, outputPath);
        arguments["data"] = new JsonArray { new JsonArray { "A" } };

        // Act & Assert - either ArgumentException from our validation or CellsException from Aspose
        await Assert.ThrowsAnyAsync<Exception>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetRange_WithMissingRange_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_no_range.xlsx");
        var arguments = CreateArguments("get", workbookPath);

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task CopyRange_WithMissingSourceRange_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_copy_no_source.xlsx");
        var outputPath = CreateTestFilePath("test_copy_no_source_output.xlsx");
        var arguments = CreateArguments("copy", workbookPath, outputPath);
        arguments["destCell"] = "B1";

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion
}