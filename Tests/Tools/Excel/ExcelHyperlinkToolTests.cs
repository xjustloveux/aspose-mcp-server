using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.Hyperlink;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelHyperlinkTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelHyperlinkToolTests : ExcelTestBase
{
    private readonly ExcelHyperlinkTool _tool;

    public ExcelHyperlinkToolTests()
    {
        _tool = new ExcelHyperlinkTool(SessionManager);
    }

    private string CreateWorkbookWithHyperlink(string fileName, string cell = "A1", string url = "https://test.com")
    {
        var workbookPath = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        workbook.Worksheets[0].Hyperlinks.Add(cell, 1, 1, url);
        workbook.Save(workbookPath);
        return workbookPath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddHyperlink()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, cell: "A1",
            url: "https://example.com", displayText: "Click here", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Hyperlink added to A1", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Hyperlinks);
        Assert.Equal("https://example.com", workbook.Worksheets[0].Hyperlinks[0].Address);
    }

    [Fact]
    public void Edit_ByCell_ShouldModifyHyperlink()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_edit_cell.xlsx", "A1", "https://old.com");
        var outputPath = CreateTestFilePath("test_edit_cell_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, cell: "A1",
            url: "https://new.com", displayText: "New Link", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Hyperlink at", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("https://new.com", workbook.Worksheets[0].Hyperlinks[0].Address);
    }

    [Fact]
    public void Delete_ByCell_ShouldDeleteHyperlink()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_delete_cell.xlsx");
        var outputPath = CreateTestFilePath("test_delete_cell_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, cell: "A1", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public void Get_ShouldReturnAllHyperlinks()
    {
        var workbookPath = CreateExcelWorkbook("test_get.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://test1.com");
            wb.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://test2.com");
            wb.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetHyperlinksExcelResult>(result);
        Assert.Equal(2, data.Count);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "A1",
            url: "https://test.com", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Hyperlink added", data.Message);
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
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, cell: "A1",
            url: "https://session.com", displayText: "Session Link");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Hyperlink added to A1", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Single(workbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_session_edit.xlsx", "A1", "https://old.com");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("edit", sessionId: sessionId, cell: "A1", url: "https://new-session.com");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Hyperlink at", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("https://new-session.com", workbook.Worksheets[0].Hyperlinks[0].Address);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithHyperlink("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete", sessionId: sessionId, cell: "A1");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://test1.com");
            wb.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://test2.com");
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetHyperlinksExcelResult>(result);
        Assert.Equal(2, data.Count);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateWorkbookWithHyperlink("test_session_file.xlsx", "A1", "https://session.com");
        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId);
        var data = GetResultData<GetHyperlinksExcelResult>(result);
        Assert.Equal(1, data.Count);
        Assert.Equal("https://session.com", data.Items.First().Url);
    }

    #endregion
}
