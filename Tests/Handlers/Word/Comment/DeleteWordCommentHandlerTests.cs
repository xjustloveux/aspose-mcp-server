using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Comment;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Comment;

public class DeleteWordCommentHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordCommentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
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

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesComment()
    {
        var doc = CreateDocumentWithComments(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted successfully", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsDeletedIndex()
    {
        var doc = CreateDocumentWithComments(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("#1", result.Message);
    }

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var doc = CreateDocumentWithComments(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Remaining comments", result.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesAtVariousIndices(int index)
    {
        var doc = CreateDocumentWithComments(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", index }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted successfully", result.Message);
    }

    [Fact]
    public void Execute_ReducesCommentCount()
    {
        var doc = CreateDocumentWithComments(3);
        var initialCount = doc.GetChildNodes(NodeType.Comment, true).Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 0 }
        });

        _handler.Execute(context, parameters);

        var finalCount = doc.GetChildNodes(NodeType.Comment, true).Count;
        Assert.Equal(initialCount - 1, finalCount);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCommentIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithComments(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("commentIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidCommentIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithComments(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeCommentIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithComments(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "commentIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
