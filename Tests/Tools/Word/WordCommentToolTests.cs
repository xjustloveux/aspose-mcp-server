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

    #region General

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
    public void AddComment_WithAuthorInitials_ShouldUseProvidedInitials()
    {
        var docPath = CreateWordDocumentWithContent("test_add_initials.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_initials_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Comment with initials", author: "Test Author", authorInitial: "XY", paragraphIndex: 0);
        var doc = new Document(outputPath);
        var comments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0);
        Assert.Equal("XY", comments[0].Initial);
    }

    [Fact]
    public void AddComment_WithNegativeOneParagraphIndex_ShouldAddToLastParagraph()
    {
        var docPath = CreateWordDocumentWithContent("test_add_last.docx", "First paragraph");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Second paragraph");
        builder.Writeln("Last paragraph");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_last_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Comment on last", author: "Author", paragraphIndex: -1);
        var resultDoc = new Document(outputPath);
        var comments = resultDoc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0);
    }

    [Fact]
    public void AddComment_WithoutParagraphIndex_ShouldAddAtEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_add_end.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_end_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, text: "Comment at end", author: "Author");
        var doc = new Document(outputPath);
        var comments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0);
    }

    [Fact]
    public void AddComment_WithRunIndexes_ShouldAddToSpecificRuns()
    {
        var docPath = CreateWordDocumentWithContent("test_add_runs.docx", "Test content for runs");
        var outputPath = CreateTestFilePath("test_add_runs_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Comment on runs", author: "Author", paragraphIndex: 0, startRunIndex: 0, endRunIndex: 0);
        var doc = new Document(outputPath);
        var comments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0);
    }

    [Fact]
    public void AddComment_WithOnlyStartRunIndex_ShouldUseAsSingleRun()
    {
        var docPath = CreateWordDocumentWithContent("test_add_single_run.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_single_run_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Single run comment", author: "Author", paragraphIndex: 0, startRunIndex: 0);
        var doc = new Document(outputPath);
        var comments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0);
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
    public void GetComments_ShouldReturnJsonWithCorrectFields()
    {
        var docPath = CreateWordDocumentWithContent("test_get_fields.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Test Author", "TA", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Test comment"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.Contains("\"author\"", result);
        Assert.Contains("\"initial\"", result);
        Assert.Contains("\"date\"", result);
        Assert.Contains("\"content\"", result);
        Assert.Contains("\"replyCount\"", result);
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

    [Fact]
    public void ReplyToComment_WithTextParameter_ShouldAddReply()
    {
        var docPath = CreateWordDocumentWithContent("test_reply_text.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Author", "AU", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Comment"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_reply_text_output.docx");
        var result = _tool.Execute("reply", docPath, outputPath: outputPath,
            commentIndex: 0, text: "Reply via text param", author: "Author");
        Assert.StartsWith("Reply added", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            text: "Case test comment", author: "Author", paragraphIndex: 0);
        Assert.StartsWith("Comment added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_get_{operation}.docx", "Test");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("count", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_del_{operation}.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Author", "AU", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Comment"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_case_del_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, commentIndex: 0);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("REPLY")]
    [InlineData("Reply")]
    [InlineData("reply")]
    public void Operation_ShouldBeCaseInsensitive_Reply(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_reply_{operation}.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Author", "AU", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Comment"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_case_reply_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            commentIndex: 0, replyText: "Reply", author: "Author");
        Assert.StartsWith("Reply added", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void AddComment_WithEmptyOrNullText_ShouldThrowArgumentException(string? text)
    {
        var docPath = CreateWordDocumentWithContent("test_add_empty_text.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_empty_text_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: text, author: "Author"));
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
    public void AddComment_WithNegativeStartRunIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_neg_run.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_neg_run_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "Comment",
                paragraphIndex: 0, startRunIndex: -1, endRunIndex: 0));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void AddComment_WithStartGreaterThanEnd_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_start_gt_end.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_start_gt_end_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "Comment",
                paragraphIndex: 0, startRunIndex: 1, endRunIndex: 0));
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

    #endregion

    #region Session

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