using System.Text.Json;
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

    private string CreateWorkbookWithCellValue(string fileName, string cell = "A1", object? value = null)
    {
        var filePath = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(filePath);
        workbook.Worksheets[0].Cells[cell].Value = value ?? "TestValue";
        workbook.Save(filePath);
        return filePath;
    }

    private string CreateWorkbookWithFormula(string fileName)
    {
        var filePath = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(filePath);
        var ws = workbook.Worksheets[0];
        ws.Cells["A1"].Value = 10;
        ws.Cells["B1"].Value = 20;
        ws.Cells["C1"].Formula = "=A1+B1";
        workbook.Save(filePath);
        return filePath;
    }

    private string CreateWorkbookWithFormattedCell(string fileName, string cell = "A1")
    {
        var filePath = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(filePath);
        var cellObj = workbook.Worksheets[0].Cells[cell];
        cellObj.Value = "FormattedValue";
        var style = cellObj.GetStyle();
        style.Font.IsBold = true;
        cellObj.SetStyle(style);
        workbook.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Write_ShouldWriteValue()
    {
        var workbookPath = CreateExcelWorkbook("test_write.xlsx");
        var outputPath = CreateTestFilePath("test_write_output.xlsx");
        var result = _tool.Execute("write", workbookPath, cell: "A1", value: "Test Value", outputPath: outputPath);
        Assert.Contains("written", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Test Value", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Theory]
    [InlineData("123.45")]
    [InlineData("true")]
    [InlineData("2024-01-15")]
    [InlineData("Hello World")]
    public void Write_WithDifferentDataTypes_ShouldWriteCorrectly(string value)
    {
        var workbookPath = CreateExcelWorkbook($"test_write_{value.Replace(".", "_")}.xlsx");
        var outputPath = CreateTestFilePath($"test_write_{value.Replace(".", "_")}_output.xlsx");
        _tool.Execute("write", workbookPath, cell: "A1", value: value, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.NotNull(workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Write_WithDateValue_ShouldWriteAsDate()
    {
        var workbookPath = CreateExcelWorkbook("test_write_date.xlsx");
        var outputPath = CreateTestFilePath("test_write_date_output.xlsx");
        _tool.Execute("write", workbookPath, cell: "A1", value: "2024-01-15", outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var dateValue = workbook.Worksheets[0].Cells["A1"].DateTimeValue;
        Assert.Equal(new DateTime(2024, 1, 15), dateValue.Date);
    }

    [Fact]
    public void Get_ShouldReturnValue()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_get.xlsx", "A1", "TestData");
        var result = _tool.Execute("get", workbookPath, cell: "A1");
        var json = JsonDocument.Parse(result);
        Assert.Equal("TestData", json.RootElement.GetProperty("value").GetString());
    }

    [Fact]
    public void Get_WithIncludeFormat_ShouldReturnFormatInfo()
    {
        var workbookPath = CreateWorkbookWithFormattedCell("test_get_format.xlsx");
        var result = _tool.Execute("get", workbookPath, cell: "A1", includeFormat: true);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("format", out var format));
        Assert.True(format.GetProperty("bold").GetBoolean());
    }

    [Fact]
    public void Get_WithCalculateFormula_ShouldReturnCalculatedValue()
    {
        var workbookPath = CreateWorkbookWithFormula("test_get_calc.xlsx");
        var result = _tool.Execute("get", workbookPath, cell: "C1", calculateFormula: true);
        var json = JsonDocument.Parse(result);
        Assert.Equal("30", json.RootElement.GetProperty("value").GetString());
    }

    [Fact]
    public void Get_WithIncludeFormula_ShouldReturnFormula()
    {
        var workbookPath = CreateWorkbookWithFormula("test_get_formula.xlsx");
        var result = _tool.Execute("get", workbookPath, cell: "C1", includeFormula: true);
        var json = JsonDocument.Parse(result);
        Assert.Equal("=A1+B1", json.RootElement.GetProperty("formula").GetString());
    }

    [SkippableFact]
    public void Get_FromDifferentSheet_ShouldGetFromCorrectSheet()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Multiple sheet operations may be limited in evaluation mode");
        var workbookPath = CreateExcelWorkbook("test_get_sheet.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Sheet1Data";
            var sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].Value = "Sheet2Data";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath, cell: "A1", sheetIndex: 1);
        Assert.Contains("Sheet2Data", result);
    }

    [Fact]
    public void Edit_WithValue_ShouldUpdateValue()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_edit.xlsx", "A1", "Original");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, cell: "A1", value: "Updated", outputPath: outputPath);
        Assert.StartsWith("Cell A1 edited", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Updated", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Edit_WithFormula_ShouldSetFormula()
    {
        var workbookPath = CreateWorkbookWithFormula("test_edit_formula.xlsx");
        var outputPath = CreateTestFilePath("test_edit_formula_output.xlsx");
        _tool.Execute("edit", workbookPath, cell: "D1", formula: "A1*B1", outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("=A1*B1", workbook.Worksheets[0].Cells["D1"].Formula);
    }

    [Fact]
    public void Edit_WithClearValue_ShouldClearCell()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_edit_clear.xlsx");
        var outputPath = CreateTestFilePath("test_edit_clear_output.xlsx");
        _tool.Execute("edit", workbookPath, cell: "A1", clearValue: true, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var value = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(value == null || value.ToString() == "");
    }

    [Fact]
    public void Clear_WithClearContent_ShouldClearContent()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_clear_content.xlsx");
        var outputPath = CreateTestFilePath("test_clear_content_output.xlsx");
        var result = _tool.Execute("clear", workbookPath, cell: "A1", clearContent: true, outputPath: outputPath);
        Assert.StartsWith("Cell A1 cleared", result);
        using var workbook = new Workbook(outputPath);
        var value = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(value == null || value.ToString() == "");
    }

    [Fact]
    public void Clear_WithClearFormat_ShouldClearFormat()
    {
        var workbookPath = CreateWorkbookWithFormattedCell("test_clear_format.xlsx");
        var outputPath = CreateTestFilePath("test_clear_format_output.xlsx");
        _tool.Execute("clear", workbookPath, cell: "A1", clearContent: false, clearFormat: true,
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.False(style.Font.IsBold);
    }

    [Fact]
    public void Clear_WithBothOptions_ShouldClearContentAndFormat()
    {
        var workbookPath = CreateWorkbookWithFormattedCell("test_clear_both.xlsx");
        var outputPath = CreateTestFilePath("test_clear_both_output.xlsx");
        _tool.Execute("clear", workbookPath, cell: "A1", clearContent: true, clearFormat: true, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var cell = workbook.Worksheets[0].Cells["A1"];
        var value = cell.Value;
        Assert.True(value == null || value.ToString() == "");
        Assert.False(cell.GetStyle().Font.IsBold);
    }

    [Theory]
    [InlineData("WRITE")]
    [InlineData("Write")]
    [InlineData("write")]
    public void Operation_ShouldBeCaseInsensitive_Write(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "A1", value: "Test", outputPath: outputPath);
        Assert.Contains("written", result); // Verify action was completed
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateWorkbookWithCellValue($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "A1");
        Assert.Contains("value", result);
    }

    [Theory]
    [InlineData("CLEAR")]
    [InlineData("Clear")]
    [InlineData("clear")]
    public void Operation_ShouldBeCaseInsensitive_Clear(string operation)
    {
        var workbookPath = CreateWorkbookWithCellValue($"test_case_clear_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_clear_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "A1", clearContent: true, outputPath: outputPath);
        Assert.Contains("cleared", result); // Verify action was completed
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath, cell: "A1"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Execute_WithEmptyOrNullCell_ShouldThrowArgumentException(string? cell)
    {
        var workbookPath = CreateExcelWorkbook("test_empty_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", workbookPath, cell: cell));
        Assert.Contains("cell is required", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidCellAddress_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", workbookPath, cell: "InvalidCell"));
        Assert.Contains("Invalid cell address format", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Write_WithEmptyOrNullValue_ShouldThrowArgumentException(string? value)
    {
        var workbookPath = CreateExcelWorkbook("test_write_no_value.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("write", workbookPath, cell: "A1", value: value));
        Assert.Contains("value is required", ex.Message);
    }

    [Fact]
    public void Edit_WithNoChanges_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_no_changes.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("edit", workbookPath, cell: "A1"));
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
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", workbookPath, cell: "A1", sheetIndex: -1));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", "", cell: "A1"));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get", cell: "A1"));
    }

    #endregion

    #region Session

    [Fact]
    public void Write_WithSessionId_ShouldWriteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_write.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("write", sessionId: sessionId, cell: "A1", value: "Session Value");
        Assert.Contains("written", result); // Verify action was completed
        Assert.Contains("session", result); // Verify session was used
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Session Value", workbook.Worksheets[0].Cells["A1"].Value?.ToString());
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_session_get.xlsx", "A1", "Session Data");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId, cell: "A1");
        Assert.Contains("Session Data", result);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_session_edit.xlsx", "A1", "Original");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, cell: "A1", value: "Updated");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Updated", workbook.Worksheets[0].Cells["A1"].Value?.ToString());
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_session_clear.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("clear", sessionId: sessionId, cell: "A1", clearContent: true);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var value = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(value == null || value.ToString() == "");
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session", cell: "A1"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateWorkbookWithCellValue("test_path_file.xlsx", "A1", "PathData");
        var sessionWorkbook = CreateWorkbookWithCellValue("test_session_file.xlsx", "A1", "SessionData");
        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId, cell: "A1");
        Assert.Contains("SessionData", result);
    }

    #endregion
}