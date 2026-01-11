using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Reference;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Reference;

public class AddCrossReferenceWordHandlerTests : WordHandlerTestBase
{
    private readonly AddCrossReferenceWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddCrossReference()
    {
        Assert.Equal("add_cross_reference", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithBookmark()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Bookmarked content");
        builder.EndBookmark("MyBookmark");
        return doc;
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsCrossReference()
    {
        var doc = CreateDocumentWithBookmark();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "Bookmark" },
            { "targetName", "MyBookmark" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("cross-reference added", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithReferenceText_AddsCrossReferenceWithText()
    {
        var doc = CreateDocumentWithBookmark();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "Heading" },
            { "targetName", "Chapter1" },
            { "referenceText", "See " }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("heading", result.ToLower());
    }

    [Fact]
    public void Execute_WithIncludeAboveBelow_AddsCrossReferenceWithPosition()
    {
        var doc = CreateDocumentWithBookmark();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "Figure" },
            { "targetName", "Figure1" },
            { "includeAboveBelow", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutReferenceType_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetName", "MyBookmark" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutTargetName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "Bookmark" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidReferenceType_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "InvalidType" },
            { "targetName", "MyBookmark" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
