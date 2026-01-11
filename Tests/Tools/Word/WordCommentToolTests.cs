using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordCommentTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordCommentToolTests : WordTestBase
{
    private readonly WordCommentTool _tool;

    public WordCommentToolTests()
    {
        _tool = new WordCommentTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddComment_ShouldAddComment()
    {
        var docPath = CreateWordDocumentWithContent("test_add_comment.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_comment_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath,
            text: "This is a test comment", author: "Test Author", paragraphIndex: 0);
        var doc = new Document(outputPath);
        var comments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0, "Document should contain at least one comment");
        Assert.Contains("test comment", comments[0].GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetComments_ShouldReturnAllComments()
    {
        var docPath = CreateWordDocumentWithContent("test_get_comments.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Test Author", "TA", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Comment", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteComment_ShouldDeleteComment()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_comment.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Test Author", "TA", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Comment to delete"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var commentsBefore = doc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.True(commentsBefore > 0);

        var outputPath = CreateTestFilePath("test_delete_comment_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, commentIndex: 0);
        var resultDoc = new Document(outputPath);
        var commentsAfter = resultDoc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.True(commentsAfter < commentsBefore);
    }

    [Fact]
    public void ReplyToComment_ShouldAddReply()
    {
        var docPath = CreateWordDocumentWithContent("test_reply_comment.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Test Author", "TA", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Original comment"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_reply_comment_output.docx");
        _tool.Execute("reply", docPath, outputPath: outputPath,
            commentIndex: 0, replyText: "This is a reply", author: "Reply Author");
        var resultDoc = new Document(outputPath);
        var comments = resultDoc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0);
        var allCommentText = string.Join(" ", comments.Select(c => c.GetText()));
        Assert.Contains("reply", allCommentText, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            text: "Case test comment", author: "Author", paragraphIndex: 0);
        Assert.StartsWith("Comment added", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddComment_WithSessionId_ShouldAddCommentInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add_comment.docx", "Test content for session");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, text: "Session comment", author: "Session Author",
            paragraphIndex: 0);
        Assert.Contains("comment", result, StringComparison.OrdinalIgnoreCase);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var comments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0);
    }

    [Fact]
    public void GetComments_WithSessionId_ShouldReturnComments()
    {
        var docPath = CreateWordDocumentWithContent("test_session_get_comments.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Session Author", "SA", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Session comment text"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("Session comment text", result);
    }

    [Fact]
    public void DeleteComment_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_delete_comment.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Author", "AU", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Comment to delete via session"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var commentsBefore = sessionDoc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.True(commentsBefore > 0);

        _tool.Execute("delete", sessionId: sessionId, commentIndex: 0);
        var commentsAfter = sessionDoc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.True(commentsAfter < commentsBefore);
    }

    [Fact]
    public void ReplyToComment_WithSessionId_ShouldAddReplyInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_reply_comment.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Original Author", "OA", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Original comment"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("reply", sessionId: sessionId,
            commentIndex: 0, replyText: "Session reply", author: "Reply Author");
        Assert.Contains("reply", result, StringComparison.OrdinalIgnoreCase);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var comments = sessionDoc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocumentWithContent("test_path_comment.docx", "Test");
        var doc1 = new Document(docPath1);
        var comment1 = new Comment(doc1, "Path Author", "PA", DateTime.Now);
        comment1.Paragraphs.Add(new Paragraph(doc1));
        comment1.FirstParagraph.Runs.Add(new Run(doc1, "PathComment"));
        doc1.FirstSection.Body.FirstParagraph.AppendChild(comment1);
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocumentWithContent("test_session_comment.docx", "Test");
        var doc2 = new Document(docPath2);
        var comment2 = new Comment(doc2, "Session Author", "SA", DateTime.Now);
        comment2.Paragraphs.Add(new Paragraph(doc2));
        comment2.FirstParagraph.Runs.Add(new Run(doc2, "SessionComment"));
        doc2.FirstSection.Body.FirstParagraph.AppendChild(comment2);
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get", docPath1, sessionId);
        Assert.Contains("SessionComment", result);
        Assert.DoesNotContain("PathComment", result);
    }

    #endregion
}
