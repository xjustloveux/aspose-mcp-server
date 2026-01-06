using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

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

    #region General

    [Fact]
    public void Add_ShouldAddComment()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, cell: "A1", comment: "This is a test comment",
            outputPath: outputPath);
        Assert.StartsWith("Comment added", result);
        using var workbook = new Workbook(outputPath);
        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.NotNull(comment);
        Assert.Contains("test comment", comment.Note, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Add_WithAuthor_ShouldSetAuthor()
    {
        var workbookPath = CreateExcelWorkbook("test_add_author.xlsx");
        var outputPath = CreateTestFilePath("test_add_author_output.xlsx");
        _tool.Execute("add", workbookPath, cell: "A1", comment: "Test comment", author: "CustomAuthor",
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.Equal("CustomAuthor", comment.Author);
    }

    [Fact]
    public void Add_WithDefaultAuthor_ShouldUseDefaultAuthor()
    {
        var workbookPath = CreateExcelWorkbook("test_add_default_author.xlsx");
        var outputPath = CreateTestFilePath("test_add_default_author_output.xlsx");
        _tool.Execute("add", workbookPath, cell: "A1", comment: "Test comment", outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.Equal("AsposeMCP", comment.Author);
    }

    [Fact]
    public void Get_ShouldReturnAllComments()
    {
        var workbookPath = CreateWorkbookWithComment("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("Test comment", result);
    }

    [Fact]
    public void Get_WithCell_ShouldReturnSpecificComment()
    {
        var workbookPath = CreateWorkbookWithComment("test_get_cell.xlsx", "B2", "Specific comment");
        var result = _tool.Execute("get", workbookPath, cell: "B2");
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("Specific comment", result);
    }

    [Fact]
    public void Get_WithNoComments_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No comments found", json.RootElement.GetProperty("message").GetString());
    }

    [Fact]
    public void Get_WithCellNoComment_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_cell_empty.xlsx");
        var result = _tool.Execute("get", workbookPath, cell: "A1");
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No comment found on cell A1", json.RootElement.GetProperty("message").GetString());
    }

    [Fact]
    public void Edit_ShouldModifyComment()
    {
        var workbookPath = CreateWorkbookWithComment("test_edit.xlsx", "A1", "Old comment");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, cell: "A1", comment: "New comment", outputPath: outputPath);
        Assert.StartsWith("Comment edited", result);
        using var workbook = new Workbook(outputPath);
        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.Contains("New", comment.Note, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Edit_WithAuthor_ShouldUpdateAuthor()
    {
        var workbookPath = CreateWorkbookWithComment("test_edit_author.xlsx", "A1", "Test comment", "OldAuthor");
        var outputPath = CreateTestFilePath("test_edit_author_output.xlsx");
        _tool.Execute("edit", workbookPath, cell: "A1", comment: "Updated", author: "NewAuthor",
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var comment = workbook.Worksheets[0].Comments["A1"];
        Assert.Equal("NewAuthor", comment.Author);
    }

    [Fact]
    public void Delete_ShouldDeleteComment()
    {
        var workbookPath = CreateWorkbookWithComment("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, cell: "A1", outputPath: outputPath);
        Assert.StartsWith("Comment deleted", result);
        using var workbook = new Workbook(outputPath);
        Assert.Null(workbook.Worksheets[0].Comments["A1"]);
    }

    [Fact]
    public void Delete_NonExistentComment_ShouldSucceed()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_nonexistent.xlsx");
        var outputPath = CreateTestFilePath("test_delete_nonexistent_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, cell: "A1", outputPath: outputPath);
        Assert.StartsWith("Comment deleted", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "A1", comment: "Test", outputPath: outputPath);
        Assert.StartsWith("Comment added", result);
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
        var workbookPath = CreateWorkbookWithComment($"test_case_delete_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_delete_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "A1", outputPath: outputPath);
        Assert.StartsWith("Comment deleted", result);
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

    [Fact]
    public void Add_WithMissingCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, comment: "Test comment"));
        Assert.Contains("cell is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingComment_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_comment.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, cell: "A1"));
        Assert.Contains("comment is required", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidCellAddress_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, cell: "InvalidCell", comment: "Test"));
        Assert.Contains("Invalid cell address format", ex.Message);
    }

    [Fact]
    public void Edit_WithMissingCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_missing_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, comment: "Test"));
        Assert.Contains("cell is required", ex.Message);
    }

    [Fact]
    public void Edit_WithMissingComment_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithComment("test_edit_missing_comment.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, cell: "A1"));
        Assert.Contains("comment is required", ex.Message);
    }

    [Fact]
    public void Edit_WithNonExistentComment_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_nonexistent.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, cell: "A1", comment: "Test"));
        Assert.Contains("No comment found", ex.Message);
    }

    [Fact]
    public void Delete_WithMissingCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_missing_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath));
        Assert.Contains("cell is required", ex.Message);
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
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, cell: "A1", comment: "Session Test Comment");
        Assert.StartsWith("Comment added", result);
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
        Assert.Contains("Session comment", result);
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
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("SessionComment", result);
    }

    #endregion
}