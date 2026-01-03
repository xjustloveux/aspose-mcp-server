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

    #region General Tests

    [Fact]
    public void AddComment_ShouldAddComment()
    {
        var workbookPath = CreateExcelWorkbook("test_add_comment.xlsx");
        var outputPath = CreateTestFilePath("test_add_comment_output.xlsx");
        _tool.Execute("add", workbookPath, cell: "A1", comment: "This is a test comment", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments["A1"];
        Assert.NotNull(comment);
        Assert.Contains("test comment", comment.Note, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetComments_ShouldReturnAllComments()
    {
        var workbookPath = CreateExcelWorkbook("test_get_comments.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments[worksheet.Comments.Add("A1")];
        comment.Note = "Test comment";
        workbook.Save(workbookPath);
        var result = _tool.Execute("get", workbookPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("note", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Test comment", result);
    }

    [Fact]
    public void EditComment_ShouldModifyComment()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_comment.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments[worksheet.Comments.Add("A1")];
        comment.Note = "Old comment";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_comment_output.xlsx");
        _tool.Execute("edit", workbookPath, cell: "A1", comment: "New comment", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var updatedComment = resultWorksheet.Comments["A1"];
        Assert.NotNull(updatedComment);
        Assert.Contains("New", updatedComment.Note, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteComment_ShouldDeleteComment()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_comment.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Comments.Add("A1");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_delete_comment_output.xlsx");
        _tool.Execute("delete", workbookPath, cell: "A1", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var comment = resultWorksheet.Comments["A1"];
        Assert.Null(comment);
    }

    [Fact]
    public void Add_WithInvalidCellAddress_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_cell_comment.xlsx");
        var outputPath = CreateTestFilePath("test_invalid_cell_comment_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, cell: "InvalidCell", comment: "Test comment", outputPath: outputPath));
        Assert.Contains("Invalid cell address format", ex.Message);
    }

    [Fact]
    public void Edit_WithNonExistentComment_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_nonexistent.xlsx");
        var outputPath = CreateTestFilePath("test_edit_nonexistent_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, cell: "A1", comment: "New comment", outputPath: outputPath));
        Assert.Contains("No comment found", ex.Message);
    }

    [Fact]
    public void Add_WithDefaultAuthor_ShouldUseDefaultAuthor()
    {
        var workbookPath = CreateExcelWorkbook("test_default_author.xlsx");
        var outputPath = CreateTestFilePath("test_default_author_output.xlsx");
        _tool.Execute("add", workbookPath, cell: "A1", comment: "Test comment without author", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments["A1"];
        Assert.NotNull(comment);
        Assert.Equal("AsposeMCP", comment.Author);
    }

    [Fact]
    public void Get_WithNoComments_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_no_comments.xlsx");
        var result = _tool.Execute("get", workbookPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No comments found", result);
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
    public void Add_WithMissingCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_cell.xlsx");
        var outputPath = CreateTestFilePath("test_missing_cell_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, comment: "Test comment", outputPath: outputPath));

        Assert.Contains("cell is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingComment_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_comment.xlsx");
        var outputPath = CreateTestFilePath("test_missing_comment_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, cell: "A1", outputPath: outputPath));

        Assert.Contains("comment is required", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_comments.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments[worksheet.Comments.Add("A1")];
        comment.Note = "Session comment";
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("Session comment", result);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add_comment.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, cell: "A1", comment: "Session Test Comment");
        Assert.Contains("Comment added", result);

        // Verify in-memory workbook has the comment
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        var sessionComment = sessionWorkbook.Worksheets[0].Comments["A1"];
        Assert.NotNull(sessionComment);
        Assert.Contains("Session Test Comment", sessionComment.Note);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_edit_comment.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments[worksheet.Comments.Add("A1")];
        comment.Note = "Original comment";
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, cell: "A1", comment: "Updated comment");

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        var updatedComment = sessionWorkbook.Worksheets[0].Comments["A1"];
        Assert.Contains("Updated", updatedComment.Note);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_delete_comment.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Comments.Add("A1");
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("delete", sessionId: sessionId, cell: "A1");

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        var deletedComment = sessionWorkbook.Worksheets[0].Comments["A1"];
        Assert.Null(deletedComment);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}