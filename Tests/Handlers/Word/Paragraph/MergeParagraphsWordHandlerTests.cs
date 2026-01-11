using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class MergeParagraphsWordHandlerTests : WordHandlerTestBase
{
    private readonly MergeParagraphsWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Merge()
    {
        Assert.Equal("merge", _handler.Operation);
    }

    #endregion

    #region Multiple Paragraphs

    [Fact]
    public void Execute_MergesThreeParagraphs()
    {
        var doc = CreateDocumentWithParagraphs("One", "Two", "Three", "Four");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("3", result);
        AssertModified(context);
    }

    #endregion

    #region Basic Merge Operations

    [Fact]
    public void Execute_MergesParagraphs()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph", "Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("merged", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_CombinesTextContent()
    {
        var doc = CreateDocumentWithParagraphs("Hello", "World", "Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 1 }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Hello");
        AssertContainsText(doc, "World");
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsMergeRange()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#0", result);
        Assert.Contains("#1", result);
    }

    [Fact]
    public void Execute_ReturnsRemainingParagraphCount()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Remaining", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutStartParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "endParagraphIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("startParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutEndParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("endParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidStartIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 100 },
            { "endParagraphIndex", 101 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Start", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidEndIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 100 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("End", ex.Message);
    }

    [Fact]
    public void Execute_WithStartGreaterThanEnd_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 2 },
            { "endParagraphIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cannot be greater", ex.Message);
    }

    [Fact]
    public void Execute_WithSameIndices_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("same", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
