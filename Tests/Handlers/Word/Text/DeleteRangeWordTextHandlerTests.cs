using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class DeleteRangeWordTextHandlerTests : WordHandlerTestBase
{
    private readonly DeleteRangeWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteRange()
    {
        Assert.Equal("delete_range", _handler.Operation);
    }

    #endregion

    #region Multi-Paragraph Deletion

    [Fact]
    public void Execute_AcrossParagraphs_DeletesRange()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph", "Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 5 },
            { "endParagraphIndex", 1 },
            { "endCharIndex", 5 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "First");
        AssertContainsText(doc, "Third paragraph");
        AssertModified(context);
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSectionIndex_DeletesInCorrectSection()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 5 },
            { "sectionIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "World");
        AssertModified(context);
    }

    #endregion

    #region Document State

    [Fact]
    public void Execute_PreservesContentOutsideRange()
    {
        var doc = CreateDocumentWithText("Hello World Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 6 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 12 }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Hello");
        AssertContainsText(doc, "Test");
        AssertModified(context);
    }

    #endregion

    #region Basic Delete Range Operations

    [Fact]
    public void Execute_DeletesTextRange()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 5 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "World");
        AssertModified(context);
    }

    [Theory]
    [InlineData(0, 5)]
    [InlineData(2, 8)]
    [InlineData(0, 11)]
    public void Execute_WithVariousRanges_DeletesCorrectly(int startChar, int endChar)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", startChar },
            { "endParagraphIndex", 0 },
            { "endCharIndex", endChar }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Error Handling - Missing Parameters

    [Fact]
    public void Execute_WithoutStartParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startCharIndex", 0 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 4 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("startParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutStartCharIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 4 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("startCharIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutEndParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endCharIndex", 4 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("endParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutEndCharIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endParagraphIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("endCharIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling - Invalid Indices

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    [InlineData(100, 0)]
    [InlineData(0, 100)]
    public void Execute_WithInvalidParagraphIndices_ThrowsArgumentException(int startPara, int endPara)
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", startPara },
            { "startCharIndex", 0 },
            { "endParagraphIndex", endPara },
            { "endCharIndex", 4 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startParagraphIndex", 0 },
            { "startCharIndex", 0 },
            { "endParagraphIndex", 0 },
            { "endCharIndex", 4 },
            { "sectionIndex", 100 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
