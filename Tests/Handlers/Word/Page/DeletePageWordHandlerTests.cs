using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Page;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Page;

public class DeletePageWordHandlerTests : WordHandlerTestBase
{
    private readonly DeletePageWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeletePage()
    {
        Assert.Equal("delete_page", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithMultiplePages()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content");
        return doc;
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesPage()
    {
        var doc = CreateDocumentWithMultiplePages();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted successfully", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_DeletesFirstPage()
    {
        var doc = CreateDocumentWithMultiplePages();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("page 0", result.ToLower());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithMultiplePages();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithMultiplePages();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
