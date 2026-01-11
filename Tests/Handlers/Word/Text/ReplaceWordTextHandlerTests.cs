using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class ReplaceWordTextHandlerTests : WordHandlerTestBase
{
    private readonly ReplaceWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Replace()
    {
        Assert.Equal("replace", _handler.Operation);
    }

    #endregion

    #region Case Sensitivity

    [Theory]
    [InlineData("Hello", true, true, false)]
    [InlineData("hello", false, true, true)]
    public void Execute_WithCaseSensitivity_ReplacesAccordingly(string find, bool caseSensitive,
        bool shouldContainReplaced, bool shouldRemoveAll)
    {
        var doc = CreateDocumentWithText("Hello HELLO hello");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", find },
            { "replace", "Replaced" },
            { "caseSensitive", caseSensitive }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        if (shouldContainReplaced)
            AssertContainsText(doc, "Replaced");
        if (shouldRemoveAll)
            AssertDoesNotContainText(doc, "Hello");
        AssertModified(context);
    }

    #endregion

    #region Replace in Fields

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void Execute_WithReplaceInFieldsOption_HandlesFieldsAccordingly(bool replaceInFields,
        bool shouldContainExcluded)
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", "Hello" },
            { "replace", "Hi" },
            { "replaceInFields", replaceInFields }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        if (shouldContainExcluded)
            Assert.Contains("excluded", result, StringComparison.OrdinalIgnoreCase);
        else
            Assert.DoesNotContain("excluded", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region No Match Scenarios

    [Fact]
    public void Execute_NoMatch_PreservesOriginalContent()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", "NotFound" },
            { "replace", "Replacement" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Hello World");
    }

    #endregion

    #region Basic Replace Operations

    [Fact]
    public void Execute_ReplacesTextInDocument()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", "World" },
            { "replace", "Universe" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Universe");
        AssertDoesNotContainText(doc, "World");
        AssertModified(context);
    }

    [Theory]
    [InlineData("Hello", "Hi", "Hello World", "Hi")]
    [InlineData("World", "Universe", "Hello World", "Universe")]
    [InlineData("foo", "bar", "foo bar baz", "bar")]
    public void Execute_ReplacesVariousTexts(string find, string replace, string originalText, string expectedInResult)
    {
        var doc = CreateDocumentWithText(originalText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", find },
            { "replace", replace }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, expectedInResult);
        AssertModified(context);
    }

    [Fact]
    public void Execute_MultipleMatches_ReplacesAll()
    {
        var doc = CreateDocumentWithText("Hello Hello Hello");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", "Hello" },
            { "replace", "Hi" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "Hi");
        AssertDoesNotContainText(doc, "Hello");
        AssertModified(context);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(5)]
    public void Execute_ReplacesMultipleOccurrences(int occurrences)
    {
        var text = string.Join(" ", Enumerable.Repeat("target", occurrences));
        var doc = CreateDocumentWithText(text);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", "target" },
            { "replace", "replaced" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "replaced");
        AssertDoesNotContainText(doc, "target");
        AssertModified(context);
    }

    #endregion

    #region Regex Replace

    [Fact]
    public void Execute_WithRegex_ReplacesPattern()
    {
        var doc = CreateDocumentWithText("Hello World123");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", @"\d+" },
            { "replace", "XXX" },
            { "useRegex", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "XXX");
        AssertDoesNotContainText(doc, "123");
        AssertModified(context);
    }

    [Theory]
    [InlineData(@"\d+", "NUM", "abc123def", "NUM")]
    [InlineData(@"\s+", "-", "a b c", "-")]
    [InlineData("[A-Z]+", "CAPS", "Hello WORLD", "CAPS")]
    public void Execute_WithVariousRegexPatterns_ReplacesCorrectly(string pattern, string replace, string originalText,
        string expected)
    {
        var doc = CreateDocumentWithText(originalText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", pattern },
            { "replace", replace },
            { "useRegex", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, expected);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithRegexFalse_TreatsPatternAsLiteral()
    {
        var doc = CreateDocumentWithText(@"Hello \d+ World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", @"\d+" },
            { "replace", "XXX" },
            { "useRegex", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, "XXX");
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFind_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "replace", "Universe" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("find", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutReplace_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", "World" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("replace", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyFind_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", "" },
            { "replace", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Document State

    [Fact]
    public void Execute_PreservesNonMatchingContent()
    {
        var doc = CreateDocumentWithText("Hello World and Universe");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", "World" },
            { "replace", "Planet" }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "Hello");
        AssertContainsText(doc, "and");
        AssertContainsText(doc, "Universe");
        AssertContainsText(doc, "Planet");
    }

    [Fact]
    public void Execute_ReplacesWithEmptyString()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", "World" },
            { "replace", "" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertDoesNotContainText(doc, "World");
        AssertModified(context);
    }

    #endregion

    #region Special Characters

    [Theory]
    [InlineData("Hello!", "Hi!", "Hello! World")]
    [InlineData("foo?", "bar?", "foo? bar")]
    [InlineData("(test)", "[test]", "(test) data")]
    public void Execute_WithSpecialCharacters_ReplacesCorrectly(string find, string replace, string originalText)
    {
        var doc = CreateDocumentWithText(originalText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", find },
            { "replace", replace }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, replace);
        AssertModified(context);
    }

    [Theory]
    [InlineData("Unicode: 中文", "替換")]
    [InlineData("日本語テスト", "置換済み")]
    public void Execute_WithUnicode_ReplacesCorrectly(string find, string replace)
    {
        var doc = CreateDocumentWithText(find);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "find", find },
            { "replace", replace }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("replaced", result, StringComparison.OrdinalIgnoreCase);
        AssertContainsText(doc, replace);
        AssertModified(context);
    }

    #endregion
}
