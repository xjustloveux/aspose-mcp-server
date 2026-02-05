using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.PageBreak;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelPageBreakTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelPageBreakToolTests : ExcelTestBase
{
    private readonly ExcelPageBreakTool _tool;

    public ExcelPageBreakToolTests()
    {
        _tool = new ExcelPageBreakTool(SessionManager);
    }

    private string CreateWorkbookWithHorizontalPageBreak(string fileName, int row = 5)
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        workbook.Worksheets[0].HorizontalPageBreaks.Add(row);
        workbook.Save(path);
        return path;
    }

    private string CreateWorkbookWithBothPageBreaks(string fileName)
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        workbook.Worksheets[0].HorizontalPageBreaks.Add(5);
        workbook.Worksheets[0].VerticalPageBreaks.Add(3);
        workbook.Save(path);
        return path;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddHorizontal_ShouldAddPageBreak()
    {
        var workbookPath = CreateExcelWorkbook("test_add_h.xlsx");
        var outputPath = CreateTestFilePath("test_add_h_output.xlsx");
        var result = _tool.Execute("add_horizontal", workbookPath, row: 5, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("page break", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].HorizontalPageBreaks.Count >= 1);
    }

    [Fact]
    public void AddVertical_ShouldAddPageBreak()
    {
        var workbookPath = CreateExcelWorkbook("test_add_v.xlsx");
        var outputPath = CreateTestFilePath("test_add_v_output.xlsx");
        var result = _tool.Execute("add_vertical", workbookPath, column: 3, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("page break", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].VerticalPageBreaks.Count >= 1);
    }

    [Fact]
    public void Get_ShouldReturnPageBreaks()
    {
        var workbookPath = CreateWorkbookWithBothPageBreaks("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetPageBreaksExcelResult>(result);
        Assert.True(data.Count >= 2);
        Assert.True(data.Items.Count >= 2);
    }

    [Fact]
    public void Remove_ShouldRemovePageBreak()
    {
        var workbookPath = CreateWorkbookWithHorizontalPageBreak("test_remove.xlsx");
        var outputPath = CreateTestFilePath("test_remove_output.xlsx");
        var result = _tool.Execute("remove", workbookPath, breakType: "horizontal", breakIndex: 0,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("removed", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].HorizontalPageBreaks);
    }

    [Fact]
    public void Clear_ShouldClearAllPageBreaks()
    {
        var workbookPath = CreateWorkbookWithBothPageBreaks("test_clear.xlsx");
        var outputPath = CreateTestFilePath("test_clear_output.xlsx");
        var result = _tool.Execute("clear", workbookPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("clear", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].HorizontalPageBreaks);
        Assert.Empty(workbook.Worksheets[0].VerticalPageBreaks);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD_HORIZONTAL")]
    [InlineData("Add_Horizontal")]
    [InlineData("add_horizontal")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, row: 5, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("page break", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddHorizontal_WithSession_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add_h.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add_horizontal", sessionId: sessionId, row: 5);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("page break", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].HorizontalPageBreaks.Count >= 1);
    }

    [Fact]
    public void AddVertical_WithSession_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add_v.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add_vertical", sessionId: sessionId, column: 3);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("page break", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].VerticalPageBreaks.Count >= 1);
    }

    [Fact]
    public void Get_WithSession_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithBothPageBreaks("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetPageBreaksExcelResult>(result);
        Assert.True(data.Count >= 2);
        var output = GetResultOutput<GetPageBreaksExcelResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Clear_WithSession_ShouldClearInMemory()
    {
        var workbookPath = CreateWorkbookWithBothPageBreaks("test_session_clear.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("clear", sessionId: sessionId);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].HorizontalPageBreaks);
        Assert.Empty(workbook.Worksheets[0].VerticalPageBreaks);
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
        var sessionWorkbook = CreateWorkbookWithHorizontalPageBreak("test_session_file.xlsx");
        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        var data = GetResultData<GetPageBreaksExcelResult>(result);
        Assert.True(data.Count >= 1);
    }

    #endregion
}
