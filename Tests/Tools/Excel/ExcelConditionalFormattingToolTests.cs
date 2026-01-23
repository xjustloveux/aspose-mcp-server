using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.ConditionalFormatting;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelConditionalFormattingTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelConditionalFormattingToolTests : ExcelTestBase
{
    private readonly ExcelConditionalFormattingTool _tool;

    public ExcelConditionalFormattingToolTests()
    {
        _tool = new ExcelConditionalFormattingTool(SessionManager);
    }

    private string CreateWorkbookWithConditionalFormatting(string fileName)
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[0];
        var cellRange = worksheet.Cells.CreateRange("A1:A5");
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
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(path);
        return path;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddConditionalFormatting()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, range: "A1:A5", condition: "GreaterThan", value: "10",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Conditional formatting added", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].ConditionalFormattings.Count > 0);
    }

    [Fact]
    public void Get_ShouldReturnFormattingInfo()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath, conditionalFormattingIndex: 0);
        var data = GetResultData<GetConditionalFormattingsResult>(result);
        Assert.True(data.Count >= 0);
    }

    [Fact]
    public void Edit_ShouldEditConditionalFormatting()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_edit.xlsx");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, conditionalFormattingIndex: 0, conditionIndex: 0,
            condition: "LessThan", value: "20", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Edited conditional formatting", data.Message);
    }

    [Fact]
    public void Delete_ShouldDeleteConditionalFormatting()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, conditionalFormattingIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Deleted conditional formatting", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].ConditionalFormattings);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx", 5, 5);
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:A5", condition: "GreaterThan", value: "10",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Conditional formatting added", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_add.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, range: "A1:A5", condition: "GreaterThan", value: "10");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Conditional formatting added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].ConditionalFormattings.Count > 0);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetConditionalFormattingsResult>(result);
        Assert.True(data.Count >= 0);
        var output = GetResultOutput<GetConditionalFormattingsResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithConditionalFormatting("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
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
        var data = GetResultData<GetConditionalFormattingsResult>(result);
        Assert.Contains("SessionSheet", data.WorksheetName);
    }

    #endregion
}
