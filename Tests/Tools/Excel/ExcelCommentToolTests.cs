using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.Comment;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelCommentTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelCommentToolTests : ExcelTestBase
{
    private readonly ExcelCommentTool _tool;

    public ExcelCommentToolTests()
    {
        _tool = new ExcelCommentTool(SessionManager);
    }

    private string CreateWorkbookWithComment(string fileName, string cell = "A1", string note = "Test comment",
        string author = "TestAuthor")
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[0];
        var commentIndex = worksheet.Comments.Add(cell);
        var comment = worksheet.Comments[commentIndex];
        comment.Note = note;
        comment.Author = author;
        workbook.Save(path);
        return path;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddComment()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, cell: "A1", comment: "This is a test comment",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Comment added", data.Message);
        using var workbook = new Workbook(outputPath);
        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.NotNull(comment);
        Assert.Contains("test comment", comment.Note, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Get_ShouldReturnAllComments()
    {
        var workbookPath = CreateWorkbookWithComment("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetCommentsExcelResult>(result);
        Assert.Equal(1, data.Count);
        Assert.Contains("Test comment", data.Items[0].Note);
    }

    [Fact]
    public void Edit_ShouldModifyComment()
    {
        var workbookPath = CreateWorkbookWithComment("test_edit.xlsx", "A1", "Old comment");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, cell: "A1", comment: "New comment", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Comment edited", data.Message);
        using var workbook = new Workbook(outputPath);
        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.Contains("New", comment.Note, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Delete_ShouldDeleteComment()
    {
        var workbookPath = CreateWorkbookWithComment("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, cell: "A1", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Comment deleted", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Null(workbook.Worksheets[0].Comments["A1"]);
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
        var result = _tool.Execute(operation, workbookPath, cell: "A1", comment: "Test", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Comment added", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath, cell: "A1"));
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
        var result = _tool.Execute("add", sessionId: sessionId, cell: "A1", comment: "Session Test Comment");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Comment added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.NotNull(comment);
        Assert.Contains("Session Test Comment", comment.Note);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithComment("test_session_get.xlsx", "A1", "Session comment");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetCommentsExcelResult>(result);
        Assert.Contains("Session comment", data.Items[0].Note);
        var output = GetResultOutput<GetCommentsExcelResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithComment("test_session_edit.xlsx", "A1", "Original comment");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, cell: "A1", comment: "Updated comment");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.Contains("Updated", comment.Note);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithComment("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("delete", sessionId: sessionId, cell: "A1");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Null(workbook.Worksheets[0].Comments["A1"]);
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
        var sessionWorkbook = CreateWorkbookWithComment("test_session_file.xlsx", "A1", "SessionComment");
        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        var data = GetResultData<GetCommentsExcelResult>(result);
        Assert.Equal(1, data.Count);
        Assert.Contains("SessionComment", data.Items[0].Note);
    }

    #endregion
}
