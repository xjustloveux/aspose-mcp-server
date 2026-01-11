using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Comment;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No comments found", json.RootElement.GetProperty("message").GetString());
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

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("count", out _));
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithComments(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsCommentsArray()
    {
        var doc = CreateDocumentWithComments(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(2, json.RootElement.GetProperty("comments").GetArrayLength());
    }

    #endregion
}
