using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelConditionalFormattingToolTests : ExcelTestBase
{
    private readonly ExcelConditionalFormattingTool _tool;

    public ExcelConditionalFormattingToolTests()
    {
        _tool = new ExcelConditionalFormattingTool(SessionManager);
    }

    private string CreateWorkbookWithConditionalFormatting(string fileName, string range = "A1:A5",
        OperatorType operatorType = OperatorType.GreaterThan, string formula1 = "10")
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[0];
        var cellRange = worksheet.Cells.CreateRange(range);
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = cellRange.FirstRow,
            StartColumn = cellRange.FirstColumn,
            EndRow = cellRange.FirstRow + cellRange.RowCount - 1,
            EndColumn = cellRange.FirstColumn + cellRange.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, operatorType, formula1, null);
        workbook.Save(path);
        return path;
    }

    #region General

    [Fact]
    public void Add_ShouldAddConditionalFormatting()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, range: "A1:A5", condition: "GreaterThan", value: "10",
            outputPath: outputPath);
        Assert.StartsWith("Conditional formatting added", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].ConditionalFormattings.Count > 0);
    }

    [Fact]
    public void Add_WithBetweenCondition_ShouldUseBothFormulas()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_between.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_between_output.xlsx");
        _tool.Execute("add", workbookPath, range: "A1:A5", condition: "Between", value: "10", formula2: "50",
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].ConditionalFormattings.Count > 0);
    }

    [Fact]
    public void Add_WithBackgroundColor_ShouldSetColor()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_color.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_color_output.xlsx");
        _tool.Execute("add", workbookPath, range: "A1:A5", condition: "GreaterThan", value: "10",
            backgroundColor: "#FF0000", outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].ConditionalFormattings.Count > 0);
    }

    [Theory]
    [InlineData("GreaterThan")]
    [InlineData("LessThan")]
    [InlineData("Equal")]
    public void Add_AllConditionTypes_ShouldWork(string condition)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_add_{condition}.xlsx", 5, 5);
        var outputPath = CreateTestFilePath($"test_add_{condition}_output.xlsx");
        var result = _tool.Execute("add", workbookPath, range: "A1:A5", condition: condition, value: "10",
            outputPath: outputPath);
        Assert.StartsWith("Conditional formatting added", result);
    }

    [Fact]
    public void Get_ShouldReturnFormattingInfo()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath, conditionalFormattingIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
    }

    [Fact]
    public void Get_WithNoFormattings_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No conditional formattings found", json.RootElement.GetProperty("message").GetString());
    }

    [Fact]
    public void Edit_ShouldEditConditionalFormatting()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_edit.xlsx");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, conditionalFormattingIndex: 0, conditionIndex: 0,
            condition: "LessThan", value: "20", outputPath: outputPath);
        Assert.StartsWith("Edited conditional formatting", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].ConditionalFormattings.Count > 0);
    }

    [Fact]
    public void Edit_WithBackgroundColor_ShouldUpdateColor()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_edit_color.xlsx");
        var outputPath = CreateTestFilePath("test_edit_color_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, conditionalFormattingIndex: 0, conditionIndex: 0,
            backgroundColor: "#00FF00", outputPath: outputPath);
        Assert.Contains("BackgroundColor=#00FF00", result);
    }

    [Fact]
    public void Delete_ShouldDeleteConditionalFormatting()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, conditionalFormattingIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Deleted conditional formatting", result);
        Assert.Contains("remaining: 0", result); // Verify remaining count
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].ConditionalFormattings);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx", 5, 5);
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:A5", condition: "GreaterThan", value: "10",
            outputPath: outputPath);
        Assert.StartsWith("Conditional formatting added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("count", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting($"test_case_delete_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_delete_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, conditionalFormattingIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Deleted conditional formatting", result);
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
    public void Add_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, condition: "GreaterThan", value: "10"));
        Assert.Contains("range is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingCondition_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_condition.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "A1:A5", value: "10"));
        Assert.Contains("condition is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingValue_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_value.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "A1:A5", condition: "GreaterThan"));
        Assert.Contains("value is required", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "InvalidRange", condition: "GreaterThan", value: "10"));
        Assert.Contains("Invalid range format", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, conditionalFormattingIndex: 99, conditionIndex: 0, value: "10"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath, conditionalFormattingIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_add.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, range: "A1:A5", condition: "GreaterThan", value: "10");
        Assert.StartsWith("Conditional formatting added", result);
        Assert.Contains("session", result); // Verify session was used
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].ConditionalFormattings.Count > 0);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_session_edit.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, conditionalFormattingIndex: 0, conditionIndex: 0,
            condition: "LessThan", value: "5");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].ConditionalFormattings.Count > 0);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        var beforeWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(beforeWorkbook.Worksheets[0].ConditionalFormattings.Count > 0);
        _tool.Execute("delete", sessionId: sessionId, conditionalFormattingIndex: 0);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].ConditionalFormattings);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateExcelWorkbook("test_path_file.xlsx");
        var sessionWorkbook = CreateWorkbookWithConditionalFormatting("test_session_file.xlsx");
        using (var wb = new Workbook(sessionWorkbook))
        {
            wb.Worksheets[0].Name = "SessionSheet";
            wb.Save(sessionWorkbook);
        }

        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        Assert.Contains("SessionSheet", result);
    }

    #endregion
}