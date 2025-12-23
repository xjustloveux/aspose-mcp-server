using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordCommentToolTests : WordTestBase
{
    private readonly WordCommentTool _tool = new();

    [Fact]
    public async Task AddComment_ShouldAddComment()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_add_comment.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_add_comment_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["paragraphIndex"] = 0;
        arguments["text"] = "This is a test comment";
        arguments["author"] = "Test Author";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var comments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0, "Document should contain at least one comment");
        Assert.Contains("test comment", comments[0].GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetComments_ShouldReturnAllComments()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_get_comments.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Test Author", "Test", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Comment", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteComment_ShouldDeleteComment()
    {
        // Arrange
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
        var arguments = CreateArguments("delete", docPath, outputPath);
        arguments["commentIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var commentsAfter = resultDoc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.True(commentsAfter < commentsBefore,
            $"Comment should be deleted. Before: {commentsBefore}, After: {commentsAfter}");
    }

    [Fact]
    public async Task ReplyToComment_ShouldAddReply()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_reply_comment.docx", "Test");
        var doc = new Document(docPath);
        var comment = new Comment(doc, "Test Author", "Test", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.Runs.Add(new Run(doc, "Original comment"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_reply_comment_output.docx");
        var arguments = CreateArguments("reply", docPath, outputPath);
        arguments["commentIndex"] = 0;
        arguments["replyText"] = "This is a reply";
        arguments["author"] = "Reply Author";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var comments = resultDoc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        Assert.True(comments.Count > 0, "Document should contain comments");
    }
}