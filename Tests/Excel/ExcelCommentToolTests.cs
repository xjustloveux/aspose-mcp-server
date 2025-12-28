using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelCommentToolTests : ExcelTestBase
{
    private readonly ExcelCommentTool _tool = new();

    [Fact]
    public async Task AddComment_ShouldAddComment()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_comment.xlsx");
        var outputPath = CreateTestFilePath("test_add_comment_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1",
            ["comment"] = "This is a test comment"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments["A1"];
        Assert.NotNull(comment);
        Assert.Contains("test comment", comment.Note, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetComments_ShouldReturnAllComments()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_comments.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments[worksheet.Comments.Add("A1")];
        comment.Note = "Test comment";
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("note", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Test comment", result);
    }

    [Fact]
    public async Task EditComment_ShouldModifyComment()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_comment.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments[worksheet.Comments.Add("A1")];
        comment.Note = "Old comment";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_comment_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1",
            ["comment"] = "New comment"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var updatedComment = resultWorksheet.Comments["A1"];
        Assert.NotNull(updatedComment);
        Assert.Contains("New", updatedComment.Note, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteComment_ShouldDeleteComment()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_comment.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Comments.Add("A1");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_delete_comment_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var comment = resultWorksheet.Comments["A1"];
        Assert.Null(comment);
    }

    [Fact]
    public async Task Add_WithInvalidCellAddress_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_cell_comment.xlsx");
        var outputPath = CreateTestFilePath("test_invalid_cell_comment_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "InvalidCell",
            ["comment"] = "Test comment"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid cell address format", ex.Message);
    }

    [Fact]
    public async Task Edit_WithNonExistentComment_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_nonexistent.xlsx");
        var outputPath = CreateTestFilePath("test_edit_nonexistent_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1",
            ["comment"] = "New comment"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("No comment found", ex.Message);
    }

    [Fact]
    public async Task Add_WithDefaultAuthor_ShouldUseDefaultAuthor()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_default_author.xlsx");
        var outputPath = CreateTestFilePath("test_default_author_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1",
            ["comment"] = "Test comment without author"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var comment = worksheet.Comments["A1"];
        Assert.NotNull(comment);
        Assert.Equal("AsposeMCP", comment.Author);
    }

    [Fact]
    public async Task Get_WithNoComments_ShouldReturnEmptyResult()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_no_comments.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No comments found", result);
    }
}