using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
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

    #region Regex Pattern Length Limit

    [Fact]
    public void Execute_RegexPatternOverLengthLimit_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("some text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", new string('a', 2001) },
            { "useRegex", true }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("maximum allowed length", ex.Message);
    }

    #endregion

    #region Story Type (Issue #1 self-describing index)

    [Fact]
    public void Execute_ReportsStoryTypeForBodyAndHeaderMatches()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("BodyText");
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("HdrText");

        var bodyResult = (TextSearchResult)_handler.Execute(CreateContext(doc),
            CreateParameters(new Dictionary<string, object?> { { "searchText", "BodyText" } }));
        var headerResult = (TextSearchResult)_handler.Execute(CreateContext(doc),
            CreateParameters(new Dictionary<string, object?> { { "searchText", "HdrText" } }));

        var bodyMatch = Assert.Single(bodyResult.Matches);
        Assert.Equal("Body", bodyMatch.StoryType);
        Assert.Equal(0, bodyMatch.ParagraphIndex);
        Assert.Equal(0, bodyMatch.SectionIndex);

        var headerMatch = Assert.Single(headerResult.Matches);
        Assert.Equal("Header", headerMatch.StoryType);
        Assert.Equal("Primary", headerMatch.HeaderFooterType);
        Assert.Equal(0, headerMatch.ParagraphIndex);
        Assert.Equal(1, headerMatch.DocumentOrderIndex);
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

    /// <summary>
    ///     Search matches substrings unless <c>wholeWord</c> says otherwise.
    ///     <para>
    ///         This case has been rewritten twice, and the history is the point. It was first
    ///         written as "MatchesAccordingly" for a <c>wholeWord</c> parameter that no handler or
    ///         tool declared or read — the key was dropped, the search behaved identically either
    ///         way, and the only assertion was that the count was not negative, so a test named
    ///         after a feature that did not exist passed for rounds. It was then restated as what
    ///         actually happened: the flag is accepted and ignored.
    ///     </para>
    ///     <para>
    ///         Now the flag does something, so the row that expected three matches with
    ///         <c>wholeWord: true</c> is the one that had to change. A test that describes a gap
    ///         should fail when the gap closes; this one did.
    ///     </para>
    /// </summary>
    /// <param name="wholeWord">Whether only whole-word matches count.</param>
    /// <param name="documentText">Document body to search.</param>
    /// <param name="expectedMatches">Matches the search finds.</param>
    [Theory]
    [InlineData(true, "Hello HelloWorld Hello", 2)]
    [InlineData(false, "Hello HelloWorld Hello", 3)]
    [InlineData(false, "HelloWorld", 1)]
    [InlineData(true, "HelloWorld", 0)]
    public void Execute_SearchesSubstrings_UnlessWholeWordIsAskedFor(bool wholeWord,
        string documentText, int expectedMatches)
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

        Assert.Equal(expectedMatches, result.MatchCount);
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

    #region Whole Word

    /// <summary>Runs a search and reports what it matched.</summary>
    /// <param name="text">The document's body text.</param>
    /// <param name="options">The search parameters.</param>
    /// <returns>The matched strings, in order.</returns>
    private List<string> Search(string text, Dictionary<string, object?> options)
    {
        var result = Assert.IsType<TextSearchResult>(
            _handler.Execute(CreateContext(CreateDocumentWithText(text)),
                CreateParameters(options)));

        return result.Matches.Select(match => match.Text).ToList();
    }

    [Fact]
    public void Execute_WithWholeWord_SkipsMatchesInsideALongerWord()
    {
        // `useRegex` with `\bcat\b` has always been able to do this. The parameter is for callers
        // who should not have to reach for a regex to ask for the obvious thing.
        var matches = Search("cat concatenate cats cat.", new Dictionary<string, object?>
        {
            { "searchText", "cat" },
            { "wholeWord", true }
        });

        Assert.Equal(2, matches.Count);
    }

    [Fact]
    public void Execute_WithoutWholeWord_StillMatchesInsideALongerWord()
    {
        // The default, and the behaviour every existing caller already has.
        var matches = Search("cat concatenate cats cat.", new Dictionary<string, object?>
        {
            { "searchText", "cat" }
        });

        Assert.Equal(4, matches.Count);
    }

    [Fact]
    public void Execute_WithWholeWordAndRegex_AppliesToWhereTheMatchLands()
    {
        // The pattern keeps its own meaning: the alternation is the caller's, and whole-word is
        // decided from the match's position rather than by wrapping the pattern in boundaries,
        // which would change what the alternation covers.
        var matches = Search("dog dogged cat concat", new Dictionary<string, object?>
        {
            { "searchText", "dog|cat" },
            { "useRegex", true },
            { "wholeWord", true }
        });

        Assert.Equal(["dog", "cat"], matches);
    }

    [Theory]
    [InlineData("say cat!", 1)]
    [InlineData("(cat)", 1)]
    [InlineData("cat-scan", 1)]
    [InlineData("cat_scan", 0)]
    [InlineData("cat9", 0)]
    [InlineData("9cat", 0)]
    public void Execute_WithWholeWord_TreatsLettersDigitsAndUnderscoreAsPartOfAWord(
        string text, int expected)
    {
        // The same rule a regex boundary uses, so a caller who switches between the two gets the
        // same answer. An underscore and a digit are part of a word; punctuation is not.
        var matches = Search(text, new Dictionary<string, object?>
        {
            { "searchText", "cat" },
            { "wholeWord", true }
        });

        Assert.Equal(expected, matches.Count);
    }

    [Fact]
    public void Execute_WithWholeWordAndNoMatch_TerminatesRatherThanRescanning()
    {
        // The literal scan advances past a rejected position as well as an accepted one. Advancing
        // only on a kept match would find the same rejected occurrence for ever.
        var matches = Search("concatenate concatenate", new Dictionary<string, object?>
        {
            { "searchText", "cat" },
            { "wholeWord", true }
        });

        Assert.Empty(matches);
    }

    #endregion

    #region Session Handles (L3)

    [Fact]
    public void Execute_SessionMode_EmitsHandleOnMatches()
    {
        var doc = CreateDocumentWithText("find me here");
        var context = new OperationContext<Document> { Document = doc, SessionId = "session-1" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "find" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);
        var match = Assert.Single(result.Matches);
        Assert.False(string.IsNullOrEmpty(match.Handle));
    }

    [Fact]
    public void Execute_FileMode_DoesNotEmitHandle()
    {
        var doc = CreateDocumentWithText("find me here");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "searchText", "find" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextSearchResult>(res);
        var match = Assert.Single(result.Matches);
        Assert.Null(match.Handle);
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
            Assert.Equal(0, m.ParagraphIndex);
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
