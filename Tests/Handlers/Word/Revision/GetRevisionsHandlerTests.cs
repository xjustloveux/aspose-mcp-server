using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Revision;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Revision;

public class GetRevisionsHandlerTests : WordHandlerTestBase
{
    private readonly GetRevisionsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetRevisions()
    {
        Assert.Equal("get_revisions", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithRevisions()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Original text.");
        doc.StartTrackRevisions("TestAuthor");
        builder.Write(" Added text.");
        doc.StopTrackRevisions();
        return doc;
    }

    #endregion

    #region Basic Get Revisions Operations

    [Fact]
    public void Execute_ReturnsRevisionsList()
    {
        var doc = CreateDocumentWithRevisions();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result);
        Assert.Contains("revisions", result);
    }

    [Fact]
    public void Execute_WithNoRevisions_ReturnsEmptyList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 0", result);
    }

    [Fact]
    public void Execute_ReturnsRevisionDetails()
    {
        var doc = CreateDocumentWithRevisions();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("type", result);
        Assert.Contains("author", result);
        Assert.Contains("date", result);
    }

    [Fact]
    public void Execute_ReturnsRevisionText()
    {
        var doc = CreateDocumentWithRevisions();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("text", result);
    }

    #endregion
}
