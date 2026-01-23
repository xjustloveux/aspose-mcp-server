using AsposeMcpServer.Handlers.PowerPoint.Text;
using AsposeMcpServer.Results.PowerPoint.Text;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextReplaceResult>(res);

        Assert.Equal(0, result.ReplacementCount);
    }

    #endregion

    #region Result Properties

    [Fact]
    public void Execute_ReturnsCorrectProperties()
    {
        var pres = CreatePresentationWithText("Original Value");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "findText", "Original" },
            { "replaceText", "New" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextReplaceResult>(res);

        Assert.Equal("Original", result.FindText);
        Assert.Equal("New", result.ReplaceText);
        Assert.Equal(1, result.ReplacementCount);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextReplaceResult>(res);

        Assert.Equal("World", result.FindText);
        Assert.Equal("Universe", result.ReplaceText);
        Assert.True(result.ReplacementCount > 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextReplaceResult>(res);

        Assert.Equal("Test", result.FindText);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextReplaceResult>(res);

        Assert.Equal("New", result.ReplaceText);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextReplaceResult>(res);

        Assert.True(result.ReplacementCount > 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextReplaceResult>(res);

        Assert.True(result.ReplacementCount > 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextReplaceResult>(res);

        Assert.Equal(1, result.ReplacementCount);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<TextReplaceResult>(res);

        // Default is case-insensitive, so both TEST and test should be replaced
        Assert.True(result.ReplacementCount > 0);
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
