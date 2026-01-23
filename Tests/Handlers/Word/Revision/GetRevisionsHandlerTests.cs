using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Revision;
using AsposeMcpServer.Results.Word.Revision;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRevisionsWordResult>(res);

        Assert.NotNull(result);
        Assert.True(result.Count >= 0);
        Assert.NotNull(result.Revisions);
    }

    [Fact]
    public void Execute_WithNoRevisions_ReturnsEmptyList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRevisionsWordResult>(res);

        Assert.NotNull(result);
        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void Execute_ReturnsRevisionDetails()
    {
        var doc = CreateDocumentWithRevisions();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRevisionsWordResult>(res);

        Assert.NotNull(result);
        Assert.NotEmpty(result.Revisions);
        var firstRevision = result.Revisions[0];
        Assert.NotNull(firstRevision.Type);
        Assert.NotNull(firstRevision.Author);
        Assert.NotNull(firstRevision.Date);
    }

    [Fact]
    public void Execute_ReturnsRevisionText()
    {
        var doc = CreateDocumentWithRevisions();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRevisionsWordResult>(res);

        Assert.NotNull(result);
        Assert.NotEmpty(result.Revisions);
        Assert.NotNull(result.Revisions[0].Text);
    }

    #endregion
}
