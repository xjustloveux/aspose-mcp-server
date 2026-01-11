using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class SearchWordTextHandlerTests : WordHandlerTestBase
{
    private readonly SearchWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Search()
    {
        Assert.Equal("search", _handler.Operation);
    }

    #endregion

    #region Multiple Matches

    [Theory]
    [InlineData("Hello", "Hello Hello Hello", "3")]
    [InlineData("World", "World World", "2")]
    [InlineData("a", "aaaaa", "5")]
    public void Execute_MultipleMatches_ReturnsCorrectCount(string searchText, string documentText,
        string expectedCount)
    {
        var doc = CreateDocumentWithText(documentText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", searchText }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains(expectedCount, result);
        AssertNotModified(context);
    }

    #endregion

    #region Case Sensitivity

    [Theory]
    [InlineData("hello", false, "Hello World")]
    [InlineData("HELLO", false, "Hello World")]
    [InlineData("HeLLo", false, "Hello World")]
    [InlineData("Hello", true, "Hello HELLO hello")]
    [InlineData("hello", false, "Hello HELLO hello")]
    public void Execute_WithCaseSensitivity_FindsAccordingly(string searchText, bool caseSensitive, string documentText)
    {
        var doc = CreateDocumentWithText(documentText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", searchText },
            { "caseSensitive", caseSensitive }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("found", result, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    #endregion

    #region Whole Word Match

    [Theory]
    [InlineData(true, "Hello HelloWorld Hello")]
    [InlineData(false, "HelloWorld")]
    public void Execute_WithWholeWordOption_MatchesAccordingly(bool wholeWord, string documentText)
    {
        var doc = CreateDocumentWithText(documentText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "Hello" },
            { "wholeWord", wholeWord }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("found", result, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    #endregion

    #region Empty Document

    [Fact]
    public void Execute_EmptyDocument_ReturnsNoMatches()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "Test" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0", result);
        AssertNotModified(context);
    }

    #endregion

    #region Basic Search Operations

    [Fact]
    public void Execute_FindsTextInDocument()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "World" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("found", result, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    [Theory]
    [InlineData("Hello", "Hello World")]
    [InlineData("World", "Hello World")]
    [InlineData("test", "This is a test document")]
    [InlineData("中文", "Unicode: 中文 text")]
    public void Execute_FindsVariousTexts(string searchText, string documentText)
    {
        var doc = CreateDocumentWithText(documentText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", searchText }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("found", result, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithNoMatch_ReturnsZeroMatches()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "NotFound" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0", result);
        AssertNotModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSearchText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("searchText", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptySearchText_ReturnsNoMatches()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0", result);
        AssertNotModified(context);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithText("Hello World");
        var originalText = GetDocumentText(doc);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "Hello" }
        });

        _handler.Execute(context, parameters);

        AssertNotModified(context);
        Assert.Equal(originalText, GetDocumentText(doc));
    }

    [Fact]
    public void Execute_MultipleCalls_DoNotModifyDocument()
    {
        var doc = CreateDocumentWithText("Hello World Test");
        var context = CreateContext(doc);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "searchText", "Hello" } }));
        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "searchText", "World" } }));
        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "searchText", "Test" } }));

        AssertNotModified(context);
    }

    #endregion

    #region Special Characters

    [Theory]
    [InlineData("Hello!")]
    [InlineData("test?")]
    [InlineData("(parentheses)")]
    [InlineData("[brackets]")]
    public void Execute_WithSpecialCharacters_FindsText(string searchText)
    {
        var doc = CreateDocumentWithText($"Content {searchText} more content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", searchText }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("found", result, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    [Theory]
    [InlineData("中文測試")]
    [InlineData("日本語")]
    [InlineData("한국어")]
    public void Execute_WithUnicode_FindsText(string searchText)
    {
        var doc = CreateDocumentWithText($"Content {searchText} more content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", searchText }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("found", result, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    #endregion
}
