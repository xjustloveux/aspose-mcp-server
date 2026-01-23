using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Page;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Page;

public class InsertBlankPageWordHandlerTests : WordHandlerTestBase
{
    private readonly InsertBlankPageWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertBlankPage()
    {
        Assert.Equal("insert_blank_page", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertAtPageIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
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

    #region Basic Insert Operations

    [Fact]
    public void Execute_InsertsBlankPageAtEnd()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("blank page inserted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithPageIndex_InsertsAtSpecificPosition()
    {
        var doc = CreateDocumentWithMultiplePages();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertAtPageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("blank page inserted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
