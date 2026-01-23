using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Revision;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Revision;

public class ManageRevisionHandlerTests : WordHandlerTestBase
{
    private readonly ManageRevisionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Manage()
    {
        Assert.Equal("manage", _handler.Operation);
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

    #region Basic Manage Revision Operations

    [Fact]
    public void Execute_AcceptsSpecificRevision()
    {
        var doc = CreateDocumentWithRevisions();
        var initialCount = doc.Revisions.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "revisionIndex", 0 },
            { "action", "accept" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("accepted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(doc.Revisions.Count < initialCount);
        AssertModified(context);
    }

    [Fact]
    public void Execute_RejectsSpecificRevision()
    {
        var doc = CreateDocumentWithRevisions();
        var initialCount = doc.Revisions.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "revisionIndex", 0 },
            { "action", "reject" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("rejected", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(doc.Revisions.Count < initialCount);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoRevisions_ReturnsNoRevisionsMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "revisionIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("no revisions", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithRevisions();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "revisionIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithRevisions();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "revisionIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidAction_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithRevisions();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "revisionIndex", 0 },
            { "action", "invalid" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithDefaultAction_AcceptsRevision()
    {
        var doc = CreateDocumentWithRevisions();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "revisionIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("accepted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
