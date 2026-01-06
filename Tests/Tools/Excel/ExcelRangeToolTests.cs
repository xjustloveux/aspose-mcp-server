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

    #region General

    [Fact]
    public void Write_ShouldWriteDataToRange()
    {
        var workbookPath = CreateExcelWorkbook("test_write.xlsx");
        var outputPath = CreateTestFilePath("test_write_output.xlsx");
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
    public void Write_WithObjectFormat_ShouldWriteData()
    {
        var workbookPath = CreateExcelWorkbook("test_write_object.xlsx");
        var outputPath = CreateTestFilePath("test_write_object_output.xlsx");
        var data = "[{\"cell\": \"A1\", \"value\": \"10\"}, {\"cell\": \"B2\", \"value\": \"20\"}]";
        _tool.Execute("write", workbookPath, startCell: "A1", data: data, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(10, Convert.ToDouble(worksheet.Cells["A1"].Value));
        Assert.Equal(20, Convert.ToDouble(worksheet.Cells["B2"].Value));
    }

    [Fact]
    public void Write_WithNumericValues_ShouldStoreAsNumbers()
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
    public void Edit_ShouldEditRangeData()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit.xlsx", 3);
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
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
    public void Edit_WithClearRange_ShouldClearBeforeEdit()
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

    [Fact]
    public void Get_ShouldReturnRangeData()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get.xlsx", 3);
        var result = _tool.Execute("get", workbookPath, range: "A1:B2");
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("R1C1", result);
        Assert.Contains("R1C2", result);
    }

    [Fact]
    public void Get_WithCalculateFormulas_ShouldRecalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_get_calc.xlsx");
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
    public void Get_WithIncludeFormat_ShouldReturnFormatInfo()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_format.xlsx", 1);
        var result = _tool.Execute("get", workbookPath, range: "A1", includeFormat: true);
        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("format", out var format));
        Assert.True(format.TryGetProperty("fontName", out _));
    }

    [Fact]
    public void Clear_ShouldClearRangeContent()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_clear.xlsx", 3);
        var outputPath = CreateTestFilePath("test_clear_output.xlsx");
        _tool.Execute("clear", workbookPath, range: "A1:B2", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var a1Value = worksheet.Cells["A1"].Value;
        Assert.True(a1Value == null || a1Value.ToString() == "");
    }

    [Fact]
    public void Clear_WithClearFormat_ShouldClearFormat()
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

    [Fact]
    public void Copy_ShouldCopyRangeToDestination()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_copy.xlsx", 3);
        var outputPath = CreateTestFilePath("test_copy_output.xlsx");
        _tool.Execute("copy", workbookPath, sourceRange: "A1:B2", destCell: "C1", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(worksheet.Cells["A1"].Value?.ToString() ?? "", worksheet.Cells["C1"].Value?.ToString() ?? "");
        Assert.Equal(worksheet.Cells["B1"].Value?.ToString() ?? "", worksheet.Cells["D1"].Value?.ToString() ?? "");
        Assert.Equal(worksheet.Cells["A2"].Value?.ToString() ?? "", worksheet.Cells["C2"].Value?.ToString() ?? "");
        Assert.Equal(worksheet.Cells["B2"].Value?.ToString() ?? "", worksheet.Cells["D2"].Value?.ToString() ?? "");
    }

    [Fact]
    public void Copy_WithValuesOnly_ShouldCopyValuesOnly()
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

    [Fact]
    public void Move_ShouldMoveRangeToDestination()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_move.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var sourceA1Value = originalWorkbook.Worksheets[0].Cells["A1"].Value;
        var sourceB1Value = originalWorkbook.Worksheets[0].Cells["B1"].Value;
        var sourceA2Value = originalWorkbook.Worksheets[0].Cells["A2"].Value;
        var sourceB2Value = originalWorkbook.Worksheets[0].Cells["B2"].Value;

        var outputPath = CreateTestFilePath("test_move_output.xlsx");
        _tool.Execute("move", workbookPath, sourceRange: "A1:B2", destCell: "C1", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];

        Assert.Equal(sourceA1Value, worksheet.Cells["C1"].Value);
        Assert.Equal(sourceB1Value, worksheet.Cells["D1"].Value);
        Assert.Equal(sourceA2Value, worksheet.Cells["C2"].Value);
        Assert.Equal(sourceB2Value, worksheet.Cells["D2"].Value);

        var a1 = worksheet.Cells["A1"].Value;
        var a2 = worksheet.Cells["A2"].Value;
        var b1 = worksheet.Cells["B1"].Value;
        var b2 = worksheet.Cells["B2"].Value;
        Assert.True(a1 == null || a1.ToString() == "", "A1 should be cleared after move");
        Assert.True(a2 == null || a2.ToString() == "", "A2 should be cleared after move");
        Assert.True(b1 == null || b1.ToString() == "", "B1 should be cleared after move");
        Assert.True(b2 == null || b2.ToString() == "", "B2 should be cleared after move");
    }

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
        Assert.True(destStyle.Font.IsBold);
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

    [Theory]
    [InlineData("WRITE")]
    [InlineData("Write")]
    [InlineData("write")]
    public void Operation_ShouldBeCaseInsensitive_Write(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var data = "[[\"Test\"]]";
        var result = _tool.Execute(operation, workbookPath, startCell: "A1", data: data, outputPath: outputPath);
        Assert.Contains("A1", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx", 1);
        var result = _tool.Execute(operation, workbookPath, range: "A1");
        Assert.Contains("items", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Write_WithMissingStartCell_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_write_no_start.xlsx");
        var outputPath = CreateTestFilePath("test_write_no_start_output.xlsx");
        var data = "[[\"A\"]]";
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("write", workbookPath, data: data, outputPath: outputPath));
    }

    [Fact]
    public void Write_WithMissingData_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_write_no_data.xlsx");
        var outputPath = CreateTestFilePath("test_write_no_data_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("write", workbookPath, startCell: "A1", data: "", outputPath: outputPath));
        Assert.Contains("data", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Write_WithInvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_write_invalid_sheet.xlsx");
        var outputPath = CreateTestFilePath("test_write_invalid_sheet_output.xlsx");
        var data = "[[\"A\"]]";
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("write", workbookPath, sheetIndex: 99, startCell: "A1", data: data, outputPath: outputPath));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Get_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_no_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", workbookPath));
        Assert.Contains("range is required", ex.Message);
    }

    [Fact]
    public void Get_WithInvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_invalid_sheet.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath, sheetIndex: 99, range: "A1:B2"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Copy_WithMissingSourceRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_no_source.xlsx");
        var outputPath = CreateTestFilePath("test_copy_no_source_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("copy", workbookPath, destCell: "B1", outputPath: outputPath));
        Assert.Contains("sourceRange is required", ex.Message);
    }

    [Fact]
    public void Copy_WithMissingDestCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_no_dest.xlsx");
        var outputPath = CreateTestFilePath("test_copy_no_dest_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("copy", workbookPath, sourceRange: "A1:B2", outputPath: outputPath));
        Assert.Contains("destCell is required", ex.Message);
    }

    [Fact]
    public void CopyFormat_WithMissingDestination_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_format_no_dest.xlsx");
        var outputPath = CreateTestFilePath("test_copy_format_no_dest_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("copy_format", workbookPath, range: "A1", outputPath: outputPath));
        Assert.Contains("destRange or destCell is required", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get", range: "A1"));
    }

    #endregion

    #region Session

    [Fact]
    public void Write_WithSessionId_ShouldWriteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_write.xlsx");
        var sessionId = OpenSession(workbookPath);
        var data = "[[\"SessionA\", \"SessionB\"], [\"SessionC\", \"SessionD\"]]";
        var result = _tool.Execute("write", sessionId: sessionId, startCell: "A1", data: data);
        Assert.StartsWith("Data written", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("SessionA", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("SessionB", workbook.Worksheets[0].Cells["B1"].Value);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_edit.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var data = "[[\"X\", \"Y\"]]";
        var result = _tool.Execute("edit", sessionId: sessionId, range: "A1:B1", data: data);
        Assert.StartsWith("Range A1:B1 edited", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("X", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId, range: "A1:B2");
        Assert.NotNull(result);
        Assert.Contains("R1C1", result);
        Assert.Contains("R1C2", result);
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_clear.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("clear", sessionId: sessionId, range: "A1:B2");
        Assert.StartsWith("Range A1:B2 cleared", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var a1Value = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(a1Value == null || a1Value.ToString() == "");
    }

    [Fact]
    public void Copy_WithSessionId_ShouldCopyInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_copy.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("copy", sessionId: sessionId, sourceRange: "A1:B2", destCell: "D1");
        Assert.StartsWith("Range A1:B2 copied", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var worksheet = workbook.Worksheets[0];
        var sourceA1 = worksheet.Cells["A1"].Value?.ToString() ?? "";
        var destD1 = worksheet.Cells["D1"].Value?.ToString() ?? "";
        Assert.Equal(sourceA1, destD1);
    }

    [Fact]
    public void Move_WithSessionId_ShouldMoveInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_move.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var originalValue = workbook.Worksheets[0].Cells["A1"].Value;
        var result = _tool.Execute("move", sessionId: sessionId, sourceRange: "A1:B2", destCell: "D1");
        Assert.StartsWith("Range A1:B2 moved", result);
        Assert.Equal(originalValue, workbook.Worksheets[0].Cells["D1"].Value);
        var a1 = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(a1 == null || a1.ToString() == "");
    }

    [Fact]
    public void CopyFormat_WithSessionId_ShouldCopyInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_copy_format.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var style = wb.CreateStyle();
            style.Font.IsBold = true;
            wb.Worksheets[0].Cells["A1"].SetStyle(style);
            wb.Worksheets[0].Cells["A1"].Value = "Test";
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("copy_format", sessionId: sessionId, range: "A1", destCell: "B1");
        Assert.StartsWith("Format copied", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var destStyle = workbook.Worksheets[0].Cells["B1"].GetStyle();
        Assert.True(destStyle.Font.IsBold);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session", range: "A1"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbookWithData("test_path_file.xlsx", 2);
        var workbookPath2 = CreateExcelWorkbookWithData("test_session_file.xlsx");
        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId, range: "A1:C5");
        Assert.Contains("R5C3", result);
    }

    #endregion
}