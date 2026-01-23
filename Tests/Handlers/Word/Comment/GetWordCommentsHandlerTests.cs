using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Comment;
using AsposeMcpServer.Results.Word.Comment;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Comment;

public class GetWordCommentsHandlerTests : WordHandlerTestBase
{
    private readonly GetWordCommentsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region No Comments

    [Fact]
    public void Execute_NoComments_ReturnsEmptyResult()
    {
        var doc = CreateDocumentWithText("No comments here");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsResult>(res);

        Assert.Equal(0, result.Count);
        Assert.NotNull(result.Message);
        Assert.Contains("No comments found", result.Message);
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

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsComments()
    {
        var doc = CreateDocumentWithComments(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsResult>(res);

        Assert.True(result.Count >= 0);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithComments(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsResult>(res);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void Execute_ReturnsCommentsArray()
    {
        var doc = CreateDocumentWithComments(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetCommentsResult>(res);

        Assert.Equal(2, result.Comments.Count);
    }

    #endregion
}
