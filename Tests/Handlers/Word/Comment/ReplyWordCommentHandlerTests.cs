using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Comment;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Comment;

public class ReplyWordCommentHandlerTests : WordHandlerTestBase
{
    private readonly ReplyWordCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Reply()
    {
        Assert.Equal("reply", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithComments(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text with comments");

        for (var i = 0; i < count; i++)
        {
            var comment = new Aspose.Words.Comment(doc, $"Author{i}", $"A{i}", DateTime.UtcNow);
            comment.SetText($"Comment {i}");
            builder.CurrentParagraph.AppendChild(comment);
        }

        return doc;
    }

    #endregion

    #region Basic Reply Operations

    [Fact]
    public void Execute_AddsReply()
    {
        var doc = CreateDocumentWithComments(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "replyText", "This is a reply" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Reply added", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsCommentIndex()
    {
        var doc = CreateDocumentWithComments(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 1 },
            { "replyText", "Reply text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#1", result);
    }

    [Fact]
    public void Execute_ReturnsReplyText()
    {
        var doc = CreateDocumentWithComments(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "replyText", "My reply content" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("My reply content", result);
    }

    [Fact]
    public void Execute_ReturnsOriginalAuthor()
    {
        var doc = CreateDocumentWithComments(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "replyText", "Reply text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Original author", result);
    }

    [Fact]
    public void Execute_ReturnsReplyAuthor()
    {
        var doc = CreateDocumentWithComments(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 },
            { "replyText", "Reply text" },
            { "author", "Jane Doe" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Jane Doe", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCommentIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithComments(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "replyText", "Reply text" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("commentIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutReplyText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithComments(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message.ToLower());
    }

    [Fact]
    public void Execute_WithInvalidCommentIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithComments(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 99 },
            { "replyText", "Reply text" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeCommentIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithComments(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", -1 },
            { "replyText", "Reply text" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
