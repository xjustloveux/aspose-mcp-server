using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.List;

public class DeleteWordListItemHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordListItemHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteItem()
    {
        Assert.Equal("delete_item", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesListItem()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2", "Item 3");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted successfully", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsDeletedItemIndex()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#1", result);
    }

    [Fact]
    public void Execute_ReturnsRemainingParagraphCount()
    {
        var doc = CreateDocumentWithParagraphs("Item 1", "Item 2", "Item 3");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Remaining paragraphs:", result);
    }

    [SkippableFact]
    public void Execute_RemovesContentFromDocument()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode adds watermark to text");
        var doc = CreateDocumentWithParagraphs("Keep This", "Delete This", "Also Keep");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 1 }
        });

        _handler.Execute(context, parameters);

        AssertDoesNotContainText(doc, "Delete This");
        AssertContainsText(doc, "Keep This");
        AssertContainsText(doc, "Also Keep");
    }

    #endregion

    #region Various Paragraph Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesVariousParagraphs(int index)
    {
        var doc = CreateDocumentWithParagraphs("Item 0", "Item 1", "Item 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", index }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted successfully", result);
    }

    [Fact]
    public void Execute_DeletesFirstParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#0", result);
    }

    [Fact]
    public void Execute_DeletesLastParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#2", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Item 1");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
