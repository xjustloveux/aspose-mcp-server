using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class InsertAtPositionWordTextHandlerTests : WordHandlerTestBase
{
    private readonly InsertAtPositionWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertAtPosition()
    {
        Assert.Equal("insert_at_position", _handler.Operation);
    }

    #endregion

    #region InsertBefore Option

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Execute_WithInsertBefore_InsertsCorrectly(bool insertBefore)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "Test" },
            { "insertBefore", insertBefore }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("inserted", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Test");
        AssertModified(context);
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSectionIndex_InsertsInCorrectSection()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "Test " },
            { "sectionIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("inserted", result, StringComparison.OrdinalIgnoreCase);
        var text = GetDocumentText(doc);
        Assert.StartsWith("Test Hello", text);
        AssertModified(context);
    }

    #endregion

    #region Special Characters

    [Theory]
    [InlineData("中文測試")]
    [InlineData("日本語")]
    [InlineData("Special: !@#$%")]
    public void Execute_WithSpecialCharacters_InsertsCorrectly(string text)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 5 },
            { "text", text }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, text);
        AssertModified(context);
    }

    #endregion

    #region Basic Insert Operations

    [Theory]
    [InlineData(0, 0, "Inserted")]
    [InlineData(0, 5, "Middle")]
    public void Execute_InsertsTextAtPosition(int paragraphIndex, int charIndex, string textToInsert)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", paragraphIndex },
            { "charIndex", charIndex },
            { "text", textToInsert }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("inserted", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, textToInsert);
        AssertModified(context);
    }

    [Fact]
    public void Execute_InsertsAtBeginning()
    {
        var doc = CreateDocumentWithText("World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "Hello " }
        });

        _handler.Execute(context, parameters);

        var text = GetDocumentText(doc);
        Assert.StartsWith("Hello ", text);
        AssertContainsText(doc, "World");
        AssertModified(context);
    }

    [Fact]
    public void Execute_InsertsAtEnd()
    {
        var doc = CreateDocumentWithText("Hello");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 5 },
            { "text", " World" }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Hello World");
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "charIndex", 0 },
            { "text", "Insert" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("insertParagraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutCharIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "text", "Insert" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("charIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(100)]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException(int paragraphIndex)
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertParagraphIndex", paragraphIndex },
            { "charIndex", 0 },
            { "text", "Insert" }
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
            { "insertParagraphIndex", 0 },
            { "charIndex", 0 },
            { "text", "Insert" },
            { "sectionIndex", 100 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
