using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordCommentToolTests : WordTestBase
{
    private readonly WordCommentTool _tool;

    public WordCommentToolTests()
    {
        _tool = new WordCommentTool(SessionManager);
    }

    #region General Tests

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
        var comment = new Comment(doc, "Test Author", "Test", DateTime.Now);
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
        var comment = new Comment(doc, "Test Author", "Test", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Comment to delete"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var commentsBefore = doc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.True(commentsBefore > 0, "Comment should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_comment_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, commentIndex: 0);
        var resultDoc = new Document(outputPath);
        var commentsAfter = resultDoc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.True(commentsAfter < commentsBefore,
            $"Comment should be deleted. Before: {commentsBefore}, After: {commentsAfter}");
    }

    [Fact]
    public void ReplyToComment_ShouldAddReply()
    {
        var docPath = CreateWordDocumentWithContent("test_reply_comment.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Test Author", "Test", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Original comment"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_reply_comment_output.docx");
        _tool.Execute("reply", docPath, outputPath: outputPath,
            commentIndex: 0, replyText: "This is a reply", author: "Reply Author");
        var resultDoc = new Document(outputPath);
        var comments = resultDoc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0, "Document should contain comments");
        var allCommentText = string.Join(" ", comments.Select(c => c.GetText()));
        Assert.Contains("reply", allCommentText, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddComment_WithEmptyText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_empty_text.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_empty_text_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "", author: "Author"));

        Assert.Contains("text is required", ex.Message);
    }

    [Fact]
    public void AddComment_WithNullText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_null_text.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_null_text_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: null, author: "Author"));

        Assert.Contains("text is required", ex.Message);
    }

    [Fact]
    public void AddComment_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_invalid_para.docx", "Single paragraph");
        var outputPath = CreateTestFilePath("test_add_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "Comment", paragraphIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void AddComment_WithNegativeParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_neg_para.docx", "Single paragraph");
        var outputPath = CreateTestFilePath("test_add_neg_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "Comment", paragraphIndex: -5));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void AddComment_WithInvalidRunIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_invalid_run.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_invalid_run_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "Comment",
                paragraphIndex: 0, startRunIndex: 999, endRunIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void DeleteComment_WithoutIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_no_index.docx", "Test content");
        var outputPath = CreateTestFilePath("test_delete_no_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath));

        Assert.Contains("commentIndex is required", ex.Message);
    }

    [Fact]
    public void DeleteComment_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_invalid_index.docx", "Test content");
        var outputPath = CreateTestFilePath("test_delete_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath, commentIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void DeleteComment_WithNegativeIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_neg_index.docx", "Test content");
        var outputPath = CreateTestFilePath("test_delete_neg_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath, commentIndex: -1));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Reply_WithoutCommentIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_reply_no_index.docx", "Test content");
        var outputPath = CreateTestFilePath("test_reply_no_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("reply", docPath, outputPath: outputPath, replyText: "Reply"));

        Assert.Contains("commentIndex is required", ex.Message);
    }

    [Fact]
    public void Reply_WithEmptyReplyText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_reply_empty_text.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Author", "AU", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Comment"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_reply_empty_text_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("reply", docPath, outputPath: outputPath, commentIndex: 0, replyText: ""));

        Assert.Contains("text or replyText is required", ex.Message);
    }

    [Fact]
    public void Reply_WithInvalidCommentIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_reply_invalid_index.docx", "Test content");
        var outputPath = CreateTestFilePath("test_reply_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("reply", docPath, outputPath: outputPath, commentIndex: 999, replyText: "Reply"));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetComments_WithNoComments_ShouldReturnEmptyResult()
    {
        var docPath = CreateWordDocumentWithContent("test_get_no_comments.docx", "No comments here");
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.Contains("\"count\":0", result.Replace(" ", ""));
        Assert.Contains("No comments found", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddComment_WithSessionId_ShouldAddCommentInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add_comment.docx", "Test content for session");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, text: "Session comment", author: "Session Author",
            paragraphIndex: 0);
        Assert.Contains("comment", result, StringComparison.OrdinalIgnoreCase);

        // Verify in-memory document has the comment
        var doc = SessionManager.GetDocument<Document>(sessionId);
        var comments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0, "Session document should contain at least one comment");
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

        // Verify comment exists before deletion
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var commentsBefore = sessionDoc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.True(commentsBefore > 0, "Comment should exist before deletion");
        _tool.Execute("delete", sessionId: sessionId, commentIndex: 0);

        // Assert - verify in-memory deletion
        var commentsAfter = sessionDoc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.True(commentsAfter < commentsBefore, "Comment should be deleted in session");
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

        // Verify in-memory document has the reply
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var comments = sessionDoc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0, "Session document should contain comments");
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

        // Act - provide both path and sessionId
        var result = _tool.Execute("get", docPath1, sessionId);

        // Assert - should use sessionId, returning SessionComment not PathComment
        Assert.Contains("SessionComment", result);
        Assert.DoesNotContain("PathComment", result);
    }

    #endregion
}