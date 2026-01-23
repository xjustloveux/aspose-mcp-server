using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Revision;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Revision;

public class AcceptAllRevisionsHandlerTests : WordHandlerTestBase
{
    private readonly AcceptAllRevisionsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AcceptAll()
    {
        Assert.Equal("accept_all", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithRevisions()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Original text.");
        doc.StartTrackRevisions("Author");
        builder.Write(" Added text.");
        doc.StopTrackRevisions();
        return doc;
    }

    #endregion

    #region Basic Accept All Revisions Operations

    [Fact]
    public void Execute_AcceptsAllRevisions()
    {
        var doc = CreateDocumentWithRevisions();
        Assert.True(doc.Revisions.Count > 0);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("accepted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, doc.Revisions.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoRevisions_ReturnsZeroCount()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("0", result.Message);
    }

    #endregion
}
