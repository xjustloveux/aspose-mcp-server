using AsposeMcpServer.Handlers.Word.Page;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Page;

public class AddPageBreakWordHandlerTests : WordHandlerTestBase
{
    private readonly AddPageBreakWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddPageBreak()
    {
        Assert.Equal("add_page_break", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsPageBreakAtEnd()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("page break added", result.ToLower());
        Assert.Contains("document end", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithParagraphIndex_AddsPageBreakAfterParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph", "Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("page break added", result.ToLower());
        Assert.Contains("after paragraph 1", result.ToLower());
    }

    #endregion
}
