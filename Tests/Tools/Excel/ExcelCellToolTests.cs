using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.Cell;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelCellTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

    [Fact]
    public void Write_ShouldWriteValueAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_write.xlsx");
        var outputPath = CreateTestFilePath("test_write_output.xlsx");
        var result = _tool.Execute("write", workbookPath, cell: "A1", value: "Test Value", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("written", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Test Value", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Get_ShouldReturnValueFromFile()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_get.xlsx", "A1", "TestData");
        var result = _tool.Execute("get", workbookPath, cell: "A1");
        var data = GetResultData<GetCellResult>(result);
        Assert.Equal("TestData", data.Value);
    }

    [Fact]
    public void Edit_ShouldUpdateValueAndPersistToFile()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_edit.xlsx", "A1", "Original");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, cell: "A1", value: "Updated", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Cell A1 edited", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Updated", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Clear_ShouldClearContentAndPersistToFile()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_clear_content.xlsx");
        var outputPath = CreateTestFilePath("test_clear_content_output.xlsx");
        var result = _tool.Execute("clear", workbookPath, cell: "A1", clearContent: true, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Cell A1 cleared", data.Message);
        using var workbook = new Workbook(outputPath);
        var value = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(value == null || value.ToString() == "");
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("WRITE")]
    [InlineData("Write")]
    [InlineData("write")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "A1", value: "Test", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("written", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath, cell: "A1"));
        Assert.Contains("Unknown operation", ex.Message);
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

    #region Session Management

    [Fact]
    public void Write_WithSessionId_ShouldWriteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_write.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("write", sessionId: sessionId, cell: "A1", value: "Session Value");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("written", data.Message);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Session Value", workbook.Worksheets[0].Cells["A1"].Value?.ToString());
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_session_get.xlsx", "A1", "Session Data");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId, cell: "A1");
        var data = GetResultData<GetCellResult>(result);
        Assert.Contains("Session Data", data.Value);
        var output = GetResultOutput<GetCellResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_session_edit.xlsx", "A1", "Original");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("edit", sessionId: sessionId, cell: "A1", value: "Updated");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Updated", workbook.Worksheets[0].Cells["A1"].Value?.ToString());
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var workbookPath = CreateWorkbookWithCellValue("test_session_clear.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("clear", sessionId: sessionId, cell: "A1", clearContent: true);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var value = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(value == null || value.ToString() == "");
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<GetCellResult>(result);
        Assert.Contains("SessionData", data.Value);
    }

    #endregion
}
