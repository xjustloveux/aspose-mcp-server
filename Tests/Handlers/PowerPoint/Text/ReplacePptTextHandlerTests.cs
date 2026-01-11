using AsposeMcpServer.Handlers.PowerPoint.Text;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Text;

public class ReplacePptTextHandlerTests : PptHandlerTestBase
{
    private readonly ReplacePptTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Replace()
    {
        Assert.Equal("replace", _handler.Operation);
    }

    #endregion

    #region No Match Scenarios

    [Fact]
    public void Execute_WithNoMatch_ReturnsZeroOccurrences()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "NotFound" },
            { "replaceText", "New" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0 occurrences", result);
    }

    #endregion

    #region Basic Replace Operations

    [Fact]
    public void Execute_ReplacesText()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "World" },
            { "replaceText", "Universe" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Replaced", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsFindText()
    {
        var pres = CreatePresentationWithText("Test");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Test" },
            { "replaceText", "New" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Test", result);
    }

    [Fact]
    public void Execute_ReturnsReplaceText()
    {
        var pres = CreatePresentationWithText("Old");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Old" },
            { "replaceText", "New" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("New", result);
    }

    [Fact]
    public void Execute_ReturnsOccurrenceCount()
    {
        var pres = CreatePresentationWithText("Test Test Test");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Test" },
            { "replaceText", "New" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("occurrences", result);
    }

    #endregion

    #region Match Case Parameter

    [Fact]
    public void Execute_WithMatchCaseFalse_MatchesCaseInsensitive()
    {
        var pres = CreatePresentationWithText("Hello HELLO hello");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "hello" },
            { "replaceText", "hi" },
            { "matchCase", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Replaced", result);
    }

    [Fact]
    public void Execute_WithMatchCaseTrue_MatchesCaseSensitive()
    {
        var pres = CreatePresentationWithText("Hello HELLO hello");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Hello" },
            { "replaceText", "Hi" },
            { "matchCase", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Replaced", result);
    }

    [Fact]
    public void Execute_DefaultMatchCase_IsCaseInsensitive()
    {
        var pres = CreatePresentationWithText("TEST test");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "test" },
            { "replaceText", "new" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Replaced", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFindText_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "replaceText", "New" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutReplaceText_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Text" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
