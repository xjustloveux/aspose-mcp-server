using System.Drawing;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Comment;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptCommentTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptCommentToolTests : PptTestBase
{
    private readonly PptCommentTool _tool;

    public PptCommentToolTests()
    {
        _tool = new PptCommentTool(SessionManager);
    }

    private string CreatePresentationWithComment(string fileName, string text = "Test comment",
        string authorName = "Test Author")
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var author = presentation.CommentAuthors.AddAuthor(authorName, "TA");
        author.Comments.AddComment(text, presentation.Slides[0], new PointF(0, 0), DateTime.UtcNow);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddComment()
    {
        var pptPath = CreatePresentation("test_add.pptx");
        var outputPath = CreateTestFilePath("test_add_output.pptx");
        var result = _tool.Execute("add", pptPath, text: "New comment", author: "Author", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Comment added", data.Message);
        using var presentation = new Presentation(outputPath);
        var comments = presentation.Slides[0].GetSlideComments(null);
        Assert.True(comments.Length > 0);
    }

    [Fact]
    public void Get_ShouldReturnComments()
    {
        var pptPath = CreatePresentationWithComment("test_get.pptx");
        var result = _tool.Execute("get", pptPath);
        var data = GetResultData<GetCommentsPptResult>(result);
        Assert.Equal(1, data.Count);
        Assert.Single(data.Items);
        Assert.Equal("Test comment", data.Items[0].Text);
    }

    [Fact]
    public void Delete_ShouldDeleteComment()
    {
        var pptPath = CreatePresentationWithComment("test_delete.pptx");
        var outputPath = CreateTestFilePath("test_delete_output.pptx");
        var result = _tool.Execute("delete", pptPath, commentIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message);
    }

    [Fact]
    public void Reply_ShouldAddReply()
    {
        var pptPath = CreatePresentationWithComment("test_reply.pptx");
        var outputPath = CreateTestFilePath("test_reply_output.pptx");
        var result = _tool.Execute("reply", pptPath, commentIndex: 0, text: "Reply text", author: "Replier",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Reply added", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, text: "Comment", author: "Author", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Comment added", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldReturnCommentsFromMemory()
    {
        var pptPath = CreatePresentationWithComment("test_session_get.pptx", "Session Comment");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetCommentsPptResult>(result);
        Assert.Equal(1, data.Count);
        Assert.Equal("Session Comment", data.Items[0].Text);
        var output = GetResultOutput<GetCommentsPptResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("add", sessionId: sessionId, text: "Session Comment", author: "Author");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Comment added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithComment("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("delete", sessionId: sessionId, commentIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Reply_WithSessionId_ShouldReplyInMemory()
    {
        var pptPath = CreatePresentationWithComment("test_session_reply.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("reply", sessionId: sessionId, commentIndex: 0, text: "Reply", author: "Replier");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Reply added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithComment("test_path_comment.pptx", "PathComment");
        var pptPath2 = CreatePresentationWithComment("test_session_comment.pptx", "SessionComment");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId);
        var data = GetResultData<GetCommentsPptResult>(result);
        Assert.Single(data.Items);
        Assert.Equal("SessionComment", data.Items[0].Text);
    }

    #endregion
}
