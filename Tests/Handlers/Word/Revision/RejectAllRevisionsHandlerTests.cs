using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Revision;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Revision;

public class RejectAllRevisionsHandlerTests : WordHandlerTestBase
{
    private readonly RejectAllRevisionsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_RejectAll()
    {
        Assert.Equal("reject_all", _handler.Operation);
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

    #region Basic Reject All Revisions Operations

    [Fact]
    public void Execute_RejectsAllRevisions()
    {
        var doc = CreateDocumentWithRevisions();
        Assert.True(doc.Revisions.Count > 0);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("rejected", result.Message, StringComparison.OrdinalIgnoreCase);
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
