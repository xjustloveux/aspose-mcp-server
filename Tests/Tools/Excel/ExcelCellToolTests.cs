using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelCellToolTests : ExcelTestBase
{
    private readonly ExcelCellTool _tool;

    public ExcelCellToolTests()
    {
        _tool = new ExcelCellTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void GetCellValue_ShouldReturnValue()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_cell_value.xlsx", 3);
        var result = _tool.Execute("get", workbookPath, cell: "A1");
        Assert.Contains("R1C1", result);
    }

    [Fact]
    public void SetCellValue_ShouldSetValue()
    {
        var workbookPath = CreateExcelWorkbook("test_set_cell_value.xlsx");
        var outputPath = CreateTestFilePath("test_set_cell_value_output.xlsx");
        _tool.Execute("write", workbookPath, cell: "A1", value: "Test Value", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("Test Value", worksheet.Cells["A1"].Value);
    }

    [Fact]
    public void SetCellFormula_ShouldSetFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_set_cell_formula.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Value = 20;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_set_cell_formula_output.xlsx");
        _tool.Execute("edit", workbookPath, cell: "C1", formula: "A1+B1", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        // Formula includes "=" prefix in Aspose.Cells
        Assert.Equal("=A1+B1", worksheet.Cells["C1"].Formula);
    }

    [Fact]
    public void GetCellFormat_ShouldReturnFormat()
    {
        var workbookPath = CreateExcelWorkbook("test_get_cell_format.xlsx");
        var workbook = new Workbook(workbookPath);
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Test";
        var style = cell.GetStyle();
        style.Font.IsBold = true;
        cell.SetStyle(style);
        workbook.Save(workbookPath);
        var result = _tool.Execute("get", workbookPath, cell: "A1", includeFormat: true);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void ClearCell_ShouldClearCellContent()
    {
        var workbookPath = CreateExcelWorkbook("test_clear_cell.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test Value";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_clear_cell_output.xlsx");
        _tool.Execute("clear", workbookPath, cell: "A1", clearContent: true, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        // Clearing cell sets value to empty string, not null
        var value = worksheet.Cells["A1"].Value;
        Assert.True(value == null || value.ToString() == "", $"Cell should be cleared, got: {value}");
    }

    [Fact]
    public void ClearCell_WithClearFormat_ShouldClearFormat()
    {
        var workbookPath = CreateExcelWorkbook("test_clear_cell_format.xlsx");
        var workbook = new Workbook(workbookPath);
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Test";
        var style = cell.GetStyle();
        style.Font.IsBold = true;
        cell.SetStyle(style);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_clear_cell_format_output.xlsx");
        _tool.Execute("clear", workbookPath, cell: "A1", clearContent: false, clearFormat: true,
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var resultStyle = worksheet.Cells["A1"].GetStyle();
        // Verify format was cleared - check that bold is false (default)
        Assert.False(resultStyle.Font.IsBold, "Cell format should be cleared (bold should be false)");
    }

    [Fact]
    public void ClearCell_WithClearContentAndFormat_ShouldClearBoth()
    {
        var workbookPath = CreateExcelWorkbook("test_clear_cell_both.xlsx");
        var workbook = new Workbook(workbookPath);
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Test";
        var style = cell.GetStyle();
        style.Font.IsBold = true;
        cell.SetStyle(style);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_clear_cell_both_output.xlsx");
        _tool.Execute("clear", workbookPath, cell: "A1", clearContent: true, clearFormat: true, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        // Clearing cell sets value to empty string, not null
        var value = worksheet.Cells["A1"].Value;
        Assert.True(value == null || value.ToString() == "", $"Cell should be cleared, got: {value}");
        // Verify format was also cleared
        var resultStyle = worksheet.Cells["A1"].GetStyle();
        Assert.False(resultStyle.Font.IsBold, "Cell format should be cleared (bold should be false)");
    }

    // Note: ExcelCellTool doesn't support setting cell format directly
    // Format operations would require a separate tool or direct Aspose.Cells API usage
    // This test is skipped as the operation doesn't exist in ExcelCellTool

    [Fact]
    public void Write_WithDifferentDataTypes_ShouldHandleTypes()
    {
        var workbookPath = CreateExcelWorkbook("test_data_types.xlsx");
        var outputPath = CreateTestFilePath("test_data_types_output.xlsx");

        // Test numeric value as string (tool converts to appropriate type)
        _tool.Execute("write", workbookPath, cell: "A1", value: "123.45", outputPath: outputPath);

        // Test boolean value as string
        _tool.Execute("write", outputPath, cell: "A2", value: "true", outputPath: outputPath);

        // Test date value as string
        _tool.Execute("write", outputPath, cell: "A3", value: "2024-01-15", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];

        // Verify values were written
        var numValue = worksheet.Cells["A1"].Value;
        Assert.NotNull(numValue);

        var boolValue = worksheet.Cells["A2"].Value;
        Assert.NotNull(boolValue);

        // Verify date/string
        var dateValue = worksheet.Cells["A3"].Value;
        Assert.NotNull(dateValue);
    }

    [SkippableFact]
    public void Get_FromDifferentSheet_ShouldGetFromSheet()
    {
        // Skip in evaluation mode - Aspose.Cells may limit operations across multiple sheets
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Multiple sheet operations may be limited in evaluation mode");
        var workbookPath = CreateExcelWorkbook("test_get_from_sheet.xlsx");

        var workbook = new Workbook(workbookPath);

        // Add data to first sheet
        workbook.Worksheets[0].Cells["A1"].Value = "Sheet1Data";

        // Add second sheet with data
        var sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].Value = "Sheet2Data";

        workbook.Save(workbookPath);
        var result = _tool.Execute("get", workbookPath, cell: "A1", sheetIndex: 1);
        Assert.Contains("Sheet2Data", result);
    }

    [Fact]
    public void Write_WithRange_ShouldWriteMultipleCells()
    {
        var workbookPath = CreateExcelWorkbook("test_write_range.xlsx");
        var outputPath = CreateTestFilePath("test_write_range_output.xlsx");
        _tool.Execute("write", workbookPath, cell: "A1", value: "MultiCell", outputPath: outputPath);

        // Add more cells
        _tool.Execute("write", outputPath, cell: "B1", value: "Second", outputPath: outputPath);
        _tool.Execute("write", outputPath, cell: "C1", value: "Third", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("MultiCell", worksheet.Cells["A1"].Value);
        Assert.Equal("Second", worksheet.Cells["B1"].Value);
        Assert.Equal("Third", worksheet.Cells["C1"].Value);
    }

    [Fact]
    public void Write_WithInvalidCellAddress_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_cell.xlsx");
        var outputPath = CreateTestFilePath("test_invalid_cell_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("write", workbookPath, cell: "InvalidCell", value: "Test", outputPath: outputPath));
        Assert.Contains("Invalid cell address format", ex.Message);
    }

    [Fact]
    public void Write_WithDateValue_ShouldWriteAsDate()
    {
        var workbookPath = CreateExcelWorkbook("test_date_value.xlsx");
        var outputPath = CreateTestFilePath("test_date_value_output.xlsx");
        _tool.Execute("write", workbookPath, cell: "A1", value: "2024-01-15", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var cell = worksheet.Cells["A1"];
        Assert.NotNull(cell.Value);
        // Excel stores dates as numeric values, verify using DateTimeValue
        var dateValue = cell.DateTimeValue;
        Assert.Equal(new DateTime(2024, 1, 15), dateValue.Date);
    }

    [Fact]
    public void Get_WithCalculateFormula_ShouldReturnCalculatedValue()
    {
        var workbookPath = CreateExcelWorkbook("test_calculate_formula.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Value = 20;
        workbook.Worksheets[0].Cells["C1"].Formula = "=A1+B1";
        workbook.Save(workbookPath);
        var result = _tool.Execute("get", workbookPath, cell: "C1", calculateFormula: true);
        Assert.Contains("30", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", workbookPath, cell: "A1"));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_empty_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath, cell: ""));

        Assert.Contains("cell is required", ex.Message);
    }

    [Fact]
    public void Execute_WithNullCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_null_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath, cell: null));

        Assert.Contains("cell is required", ex.Message);
    }

    [Fact]
    public void Write_WithoutValue_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_write_no_value.xlsx");
        var outputPath = CreateTestFilePath("test_write_no_value_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("write", workbookPath, cell: "A1", value: null, outputPath: outputPath));

        Assert.Contains("value is required", ex.Message);
    }

    [Fact]
    public void Write_WithEmptyValue_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_write_empty_value.xlsx");
        var outputPath = CreateTestFilePath("test_write_empty_value_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("write", workbookPath, cell: "A1", value: "", outputPath: outputPath));

        Assert.Contains("value is required", ex.Message);
    }

    [Fact]
    public void Edit_WithNoChanges_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_no_changes.xlsx");
        var outputPath = CreateTestFilePath("test_edit_no_changes_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, cell: "A1", outputPath: outputPath));

        Assert.Contains("Either value, formula, or clearValue must be provided", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath, cell: "A1", sheetIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_neg_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath, cell: "A1", sheetIndex: -1));

        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Write_WithSessionId_ShouldWriteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_write.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("write", sessionId: sessionId, cell: "A1", value: "Session Value");
        Assert.Contains("A1", result);

        // Verify in-memory workbook has the value
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Session Value", workbook.Worksheets[0].Cells["A1"].Value?.ToString());
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test Data";
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId, cell: "A1");
        Assert.Contains("Test Data", result);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_edit.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Original";
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, cell: "A1", value: "Updated");

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Updated", sessionWorkbook.Worksheets[0].Cells["A1"].Value?.ToString());
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_clear.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Clear Me";
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("clear", sessionId: sessionId, cell: "A1", clearContent: true);

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        var value = sessionWorkbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(value == null || value.ToString() == "");
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id", cell: "A1"));
    }

    #endregion
}