using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Comment;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Comment;

public class AddWordCommentHandlerTests : WordHandlerTestBase
{
    private readonly AddWordCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithParagraphs(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++) builder.Writeln($"Paragraph {i}");
        return doc;
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsComment()
    {
        var doc = CreateDocumentWithText("Sample text for comment");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "This is a comment" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Comment added successfully", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsCommentText()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "My comment content" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("My comment content", result);
    }

    [Fact]
    public void Execute_ReturnsAuthor()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" },
            { "author", "John Doe" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("John Doe", result);
    }

    [Fact]
    public void Execute_DefaultsAuthor()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Author:", result);
    }

    #endregion

    #region Paragraph Index

    [Fact]
    public void Execute_WithParagraphIndex_AddsToSpecificParagraph()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment on second paragraph" },
            { "paragraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Comment added successfully", result);
    }

    [Fact]
    public void Execute_WithParagraphIndexMinusOne_AddsToLastParagraph()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment on last paragraph" },
            { "paragraphIndex", -1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Comment added successfully", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comment text" },
            { "paragraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
