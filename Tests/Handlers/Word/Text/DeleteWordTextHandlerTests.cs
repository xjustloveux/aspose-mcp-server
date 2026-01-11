using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class DeleteWordTextHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Delete by Paragraph Index

    [Fact]
    public void Execute_WithParagraphIndices_DeletesTextRange()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph", "Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 1 },
            { "endParagraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        AssertDoesNotContainText(doc, "Second paragraph");
        AssertContainsText(doc, "First paragraph");
        AssertContainsText(doc, "Third paragraph");
        AssertModified(context);
    }

    #endregion

    #region Error Handling - Invalid Indices

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    [InlineData(100, 100)]
    public void Execute_WithInvalidIndices_ThrowsArgumentException(int startIdx, int endIdx)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", startIdx },
            { "endParagraphIndex", endIdx }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Delete by SearchText

    [Theory]
    [InlineData("Hello")]
    [InlineData("World")]
    [InlineData("Test")]
    public void Execute_WithSearchText_DeletesMatchingText(string searchText)
    {
        var doc = CreateDocumentWithText($"Before {searchText} After");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", searchText }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        AssertDoesNotContainText(doc, searchText);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSearchTextNotFound_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "NotFound" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling - Missing Parameters

    [Fact]
    public void Execute_WithoutRequiredParams_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithOnlyStartIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("endParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
