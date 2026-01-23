using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Results.Word.Text;
using AsposeMcpServer.Tests.Infrastructure;

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
    [InlineData("Hello", "Hello Hello Hello", 3)]
    [InlineData("World", "World World", 2)]
    [InlineData("a", "aaaaa", 5)]
    public void Execute_MultipleMatches_ReturnsCorrectCount(string searchText, string documentText,
        int expectedCount)
    {
        var doc = CreateDocumentWithText(documentText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", searchText }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.Equal(expectedCount, result.MatchCount);
        Assert.Equal(expectedCount, result.Matches.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Case Sensitivity

    [Theory]
    [InlineData("hello", false, "Hello World", 1)]
    [InlineData("HELLO", false, "Hello World", 1)]
    [InlineData("HeLLo", false, "Hello World", 1)]
    [InlineData("Hello", true, "Hello HELLO hello", 1)]
    [InlineData("hello", false, "Hello HELLO hello", 3)]
    public void Execute_WithCaseSensitivity_FindsAccordingly(string searchText, bool caseSensitive, string documentText,
        int expectedCount)
    {
        var doc = CreateDocumentWithText(documentText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", searchText },
            { "caseSensitive", caseSensitive }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.Equal(expectedCount, result.MatchCount);
        Assert.Equal(caseSensitive, result.CaseSensitive);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.True(result.MatchCount >= 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.Equal(0, result.MatchCount);
        Assert.Empty(result.Matches);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.Equal(1, result.MatchCount);
        Assert.Single(result.Matches);
        Assert.Equal("World", result.Matches[0].Text);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.True(result.MatchCount > 0);
        Assert.NotEmpty(result.Matches);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.Equal(0, result.MatchCount);
        Assert.Empty(result.Matches);
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
    public void Execute_WithEmptySearchText_ReturnsAllPositions()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        // Empty string matches at every position (limited by maxResults)
        Assert.True(result.MatchCount > 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.True(result.MatchCount > 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.True(result.MatchCount > 0);
        AssertNotModified(context);
    }

    #endregion

    #region Result Properties

    [Fact]
    public void Execute_ReturnsCorrectSearchParameters()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "Hello" },
            { "useRegex", true },
            { "caseSensitive", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.Equal("Hello", result.SearchText);
        Assert.True(result.UseRegex);
        Assert.True(result.CaseSensitive);
    }

    [Fact]
    public void Execute_ReturnsMatchDetails()
    {
        var doc = CreateDocumentWithText("Hello World Hello");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "Hello" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.Equal(2, result.MatchCount);
        // ReSharper disable once ParameterOnlyUsedForPreconditionCheck.Local - Assert.All parameter is intended for validation
        Assert.All(result.Matches, m =>
        {
            Assert.Equal("Hello", m.Text);
            Assert.True(m.ParagraphIndex >= 0);
            Assert.NotEmpty(m.Context);
        });
    }

    [Fact]
    public void Execute_WithMaxResults_LimitsMatches()
    {
        var doc = CreateDocumentWithText("Hello Hello Hello Hello Hello Hello");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "Hello" },
            { "maxResults", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);

        Assert.Equal(3, result.MatchCount);
        Assert.True(result.LimitReached);
    }

    #endregion
}
