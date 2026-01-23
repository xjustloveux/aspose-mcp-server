using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

public class InsertParagraphWordHandlerTests : WordHandlerTestBase
{
    private readonly InsertParagraphWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Insert()
    {
        Assert.Equal("insert", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsParagraphCountInMessage()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New paragraph" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("paragraph count", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Insert Operations

    [Fact]
    public void Execute_InsertsParagraph()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New paragraph" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "New paragraph");
        AssertModified(context);
    }

    [Theory]
    [InlineData("Hello World")]
    [InlineData("Test content")]
    [InlineData("Special chars: !@#$%")]
    public void Execute_InsertsVariousTexts(string text)
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", text }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, text);
    }

    #endregion

    #region Insert Position

    [Fact]
    public void Execute_WithParagraphIndexMinus1_InsertsAtBeginning()
    {
        var doc = CreateDocumentWithParagraphs("Existing paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New at beginning" },
            { "paragraphIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("beginning", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithParagraphIndex_InsertsAfterSpecifiedParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Inserted after first" },
            { "paragraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("after paragraph #0", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithoutParagraphIndex_InsertsAtEnd()
    {
        var doc = CreateDocumentWithParagraphs("First");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "At end" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("end of document", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "At end");
    }

    #endregion

    #region Formatting Options

    [Fact]
    public void Execute_WithStyleName_AppliesStyle()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Styled paragraph" },
            { "styleName", "Heading 1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Heading 1", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAlignment_AppliesAlignment()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Centered paragraph" },
            { "alignment", "center" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("center", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithIndentation_AppliesIndentation()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Indented paragraph" },
            { "indentLeft", 36.0 }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Indented paragraph");
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSpacing_AppliesSpacing()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Spaced paragraph" },
            { "spaceBefore", 12.0 },
            { "spaceAfter", 12.0 }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Spaced paragraph");
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New paragraph" },
            { "paragraphIndex", 100 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidStyleName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Existing");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New paragraph" },
            { "styleName", "NonExistentStyle" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Style", ex.Message);
    }

    #endregion
}
