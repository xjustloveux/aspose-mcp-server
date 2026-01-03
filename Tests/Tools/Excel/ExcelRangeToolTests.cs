using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelRangeToolTests : ExcelTestBase
{
    private readonly ExcelRangeTool _tool;

    public ExcelRangeToolTests()
    {
        _tool = new ExcelRangeTool(SessionManager);
    }

    #region General Tests

    #region Move Tests

    [Fact]
    public void MoveRange_ShouldMoveRangeToDestination()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_move_range.xlsx", 3);
        var sourceA1Value = new Workbook(workbookPath).Worksheets[0].Cells["A1"].Value;
        var outputPath = CreateTestFilePath("test_move_range_output.xlsx");
        _tool.Execute("move", workbookPath, sourceRange: "A1:B2", destCell: "C1", outputPath: outputPath);
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
    public void WriteRange_ShouldWriteDataToRange()
    {
        var workbookPath = CreateExcelWorkbook("test_write_range.xlsx");
        var outputPath = CreateTestFilePath("test_write_range_output.xlsx");
        var data = "[[\"A\", \"B\"], [\"C\", \"D\"]]";
        _tool.Execute("write", workbookPath, startCell: "A1", data: data, outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("A", worksheet.Cells["A1"].Value);
        Assert.Equal("B", worksheet.Cells["B1"].Value);
        Assert.Equal("C", worksheet.Cells["A2"].Value);
        Assert.Equal("D", worksheet.Cells["B2"].Value);
    }

    [Fact]
    public void WriteRange_WithObjectFormat_ShouldWriteData()
    {
        var workbookPath = CreateExcelWorkbook("test_write_object_format.xlsx");
        var outputPath = CreateTestFilePath("test_write_object_format_output.xlsx");
        var data = "[{\"cell\": \"A1\", \"value\": \"10\"}, {\"cell\": \"B2\", \"value\": \"20\"}]";
        _tool.Execute("write", workbookPath, startCell: "A1", data: data, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(10, Convert.ToDouble(worksheet.Cells["A1"].Value));
        Assert.Equal(20, Convert.ToDouble(worksheet.Cells["B2"].Value));
    }

    [Fact]
    public void WriteRange_WithNumericValues_ShouldStoreAsNumbers()
    {
        var workbookPath = CreateExcelWorkbook("test_write_numeric.xlsx");
        var outputPath = CreateTestFilePath("test_write_numeric_output.xlsx");
        var data = "[[\"100\", \"200.5\", \"true\"]]";
        _tool.Execute("write", workbookPath, startCell: "A1", data: data, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(100, Convert.ToDouble(worksheet.Cells["A1"].Value));
        Assert.Equal(200.5, Convert.ToDouble(worksheet.Cells["B1"].Value));
        Assert.Equal(true, worksheet.Cells["C1"].Value);
    }

    [Fact]
    public void WriteRange_WithInvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_write_invalid_sheet.xlsx");
        var outputPath = CreateTestFilePath("test_write_invalid_sheet_output.xlsx");
        var data = "[[\"A\"]]";
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("write", workbookPath, sheetIndex: 99, startCell: "A1", data: data, outputPath: outputPath));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region Get Tests

    [Fact]
    public void GetRange_ShouldReturnRangeData()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_range.xlsx", 3);
        var result = _tool.Execute("get", workbookPath, range: "A1:B2");
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("R1C1", result);
        Assert.Contains("R1C2", result);
    }

    [Fact]
    public void GetRange_WithCalculateFormulas_ShouldRecalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_get_calculate.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = 10;
            wb.Worksheets[0].Cells["A2"].Value = 20;
            wb.Worksheets[0].Cells["A3"].Formula = "=A1+A2";
            wb.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath, range: "A3", calculateFormulas: true);
        Assert.Contains("30", result);
    }

    [Fact]
    public void GetRange_WithIncludeFormat_ShouldReturnFormatInfo()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_format.xlsx", 1);
        var result = _tool.Execute("get", workbookPath, range: "A1", includeFormat: true);
        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("format", out var format));
        Assert.True(format.TryGetProperty("fontName", out _));
    }

    [Fact]
    public void GetRange_WithInvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_invalid_sheet.xlsx", 3);
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath, sheetIndex: 99, range: "A1:B2"));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    #region Clear Tests

    [Fact]
    public void ClearRange_ShouldClearRangeContent()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_clear_range.xlsx", 3);
        var outputPath = CreateTestFilePath("test_clear_range_output.xlsx");
        _tool.Execute("clear", workbookPath, range: "A1:B2", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Cleared cells should be empty
        var a1Value = worksheet.Cells["A1"].Value;
        Assert.True(a1Value == null || a1Value.ToString() == "",
            $"Cell A1 should be cleared, got: {a1Value}");
    }

    [Fact]
    public void ClearRange_WithClearFormat_ShouldClearFormat()
    {
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
        _tool.Execute("clear", workbookPath, range: "A1", clearContent: false, clearFormat: true,
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var resultStyle = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.False(resultStyle.Font.IsBold);
    }

    #endregion

    #region Copy Tests

    [Fact]
    public void CopyRange_ShouldCopyRangeToDestination()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_copy_range.xlsx", 3);
        var outputPath = CreateTestFilePath("test_copy_range_output.xlsx");
        _tool.Execute("copy", workbookPath, sourceRange: "A1:B2", destCell: "C1", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Verify data was copied
        var sourceA1 = worksheet.Cells["A1"].Value?.ToString() ?? "";
        var destC1 = worksheet.Cells["C1"].Value?.ToString() ?? "";
        Assert.Equal(sourceA1, destC1);
    }

    [Fact]
    public void CopyRange_WithValuesOnly_ShouldCopyValuesOnly()
    {
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
        _tool.Execute("copy", workbookPath, sourceRange: "A1", destCell: "B1", copyOptions: "Values",
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Test", workbook.Worksheets[0].Cells["B1"].Value);
        var destStyle = workbook.Worksheets[0].Cells["B1"].GetStyle();
        Assert.False(destStyle.Font.IsBold);
    }

    #endregion

    #region Edit Tests

    [Fact]
    public void EditRange_ShouldEditRangeData()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_range.xlsx", 3);
        var outputPath = CreateTestFilePath("test_edit_range_output.xlsx");
        var data = "[[\"X\", \"Y\"], [\"Z\", \"W\"]]";
        _tool.Execute("edit", workbookPath, range: "A1:B2", data: data, outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("X", worksheet.Cells["A1"].Value);
        Assert.Equal("Y", worksheet.Cells["B1"].Value);
        Assert.Equal("Z", worksheet.Cells["A2"].Value);
        Assert.Equal("W", worksheet.Cells["B2"].Value);
    }

    [Fact]
    public void EditRange_WithClearRange_ShouldClearBeforeEdit()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_clear.xlsx", 3);
        var outputPath = CreateTestFilePath("test_edit_clear_output.xlsx");
        var data = "[[\"X\"]]";
        _tool.Execute("edit", workbookPath, range: "A1:C3", data: data, clearRange: true, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("X", worksheet.Cells["A1"].Value);
        var b1 = worksheet.Cells["B1"].Value;
        Assert.True(b1 == null || b1.ToString() == "");
    }

    #endregion

    #region CopyFormat Tests

    [Fact]
    public void CopyFormat_ShouldCopyFormatOnly()
    {
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
        _tool.Execute("copy_format", workbookPath, sourceRange: "A1", destCell: "B1", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var destStyle = resultWorksheet.Cells["B1"].GetStyle();
        Assert.True(destStyle.Font.IsBold, "Format should be copied (bold should be true)");
        Assert.Equal(14, destStyle.Font.Size);
    }

    [Fact]
    public void CopyFormat_WithCopyValue_ShouldCopyFormatAndValues()
    {
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
        _tool.Execute("copy_format", workbookPath, range: "A1", destCell: "B1", copyValue: true,
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Original", workbook.Worksheets[0].Cells["B1"].Value);
        var destStyle = workbook.Worksheets[0].Cells["B1"].GetStyle();
        Assert.True(destStyle.Font.IsBold);
    }

    [Fact]
    public void CopyFormat_WithMissingDestination_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_format_no_dest.xlsx");
        var outputPath = CreateTestFilePath("test_copy_format_no_dest_output.xlsx");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("copy_format", workbookPath, range: "A1", outputPath: outputPath));
        Assert.Contains("destRange or destCell is required", exception.Message);
    }

    #endregion

    #region Error Handling Tests

    [Fact]
    public void ExecuteAsync_WithMissingPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", ""));
    }

    [Fact]
    public void WriteRange_WithMissingStartCell_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_write_no_start.xlsx");
        var outputPath = CreateTestFilePath("test_write_no_start_output.xlsx");
        var data = "[[\"A\"]]";

        // Act & Assert - either ArgumentException from our validation or CellsException from Aspose
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("write", workbookPath, data: data, outputPath: outputPath));
    }

    [Fact]
    public void GetRange_WithMissingRange_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_no_range.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath));
    }

    [Fact]
    public void CopyRange_WithMissingSourceRange_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_no_source.xlsx");
        var outputPath = CreateTestFilePath("test_copy_no_source_output.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("copy", workbookPath, destCell: "B1", outputPath: outputPath));
    }

    #endregion

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_operation.xlsx");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("invalid_operation", workbookPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void Write_WithMissingData_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_write_no_data.xlsx");
        var outputPath = CreateTestFilePath("test_write_no_data_output.xlsx");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("write", workbookPath, startCell: "A1", data: "", outputPath: outputPath));
        Assert.Contains("data", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetRange_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get_range.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId, range: "A1:B2");
        Assert.NotNull(result);
        Assert.Contains("R1C1", result);
        Assert.Contains("R1C2", result);
    }

    [Fact]
    public void WriteRange_WithSessionId_ShouldWriteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_write_range.xlsx");
        var sessionId = OpenSession(workbookPath);
        var data = "[[\"SessionA\", \"SessionB\"], [\"SessionC\", \"SessionD\"]]";
        _tool.Execute("write", sessionId: sessionId, startCell: "A1", data: data);

        // Assert - verify in-memory change
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("SessionA", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("SessionB", workbook.Worksheets[0].Cells["B1"].Value);
        Assert.Equal("SessionC", workbook.Worksheets[0].Cells["A2"].Value);
        Assert.Equal("SessionD", workbook.Worksheets[0].Cells["B2"].Value);
    }

    [Fact]
    public void CopyRange_WithSessionId_ShouldCopyInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_copy_range.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("copy", sessionId: sessionId, sourceRange: "A1:B2", destCell: "D1");

        // Assert - verify in-memory change
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var worksheet = workbook.Worksheets[0];
        var sourceA1 = worksheet.Cells["A1"].Value?.ToString() ?? "";
        var destD1 = worksheet.Cells["D1"].Value?.ToString() ?? "";
        Assert.Equal(sourceA1, destD1);
    }

    [Fact]
    public void ClearRange_WithSessionId_ShouldClearInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_clear_range.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("clear", sessionId: sessionId, range: "A1:B2");

        // Assert - verify in-memory change
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var a1Value = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(a1Value == null || a1Value.ToString() == "");
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id", range: "A1"));
    }

    #endregion
}