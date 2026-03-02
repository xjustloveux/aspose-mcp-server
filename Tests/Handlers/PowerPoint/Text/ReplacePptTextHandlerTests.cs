using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Text;
using AsposeMcpServer.Results.PowerPoint.Text;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Text;

[SupportedOSPlatform("windows")]
public class ReplacePptTextHandlerTests : PptHandlerTestBase
{
    private readonly ReplacePptTextHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Replace()
    {
        SkipIfNotWindows();
        Assert.Equal("replace", _handler.Operation);
    }

    #endregion

    #region No Match Scenarios

    [SkippableFact]
    public void Execute_WithNoMatch_ReturnsZeroOccurrences()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_ReturnsCorrectProperties()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Text replacement has limitations in evaluation mode");

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

    [SkippableFact]
    public void Execute_ReplacesText()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Text replacement has limitations in evaluation mode");

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

    [SkippableFact]
    public void Execute_ReturnsFindText()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_ReturnsReplaceText()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_ReturnsOccurrenceCount()
    {
        SkipIfNotWindows();
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

        if (!IsEvaluationMode())
        {
            var shape = pres.Slides[0].Shapes[0] as IAutoShape;
            Assert.NotNull(shape);
            Assert.Contains("New", shape.TextFrame.Text);
            Assert.DoesNotContain("Test", shape.TextFrame.Text);
        }
    }

    #endregion

    #region Match Case Parameter

    [SkippableFact]
    public void Execute_WithMatchCaseFalse_MatchesCaseInsensitive()
    {
        SkipIfNotWindows();
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

        if (!IsEvaluationMode())
        {
            var shape = pres.Slides[0].Shapes[0] as IAutoShape;
            Assert.NotNull(shape);
            Assert.Contains("hi", shape.TextFrame.Text);
        }
    }

    [SkippableFact]
    public void Execute_WithMatchCaseTrue_MatchesCaseSensitive()
    {
        SkipIfNotWindows();
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

        if (!IsEvaluationMode())
        {
            var shape = pres.Slides[0].Shapes[0] as IAutoShape;
            Assert.NotNull(shape);
            Assert.Contains("Hi", shape.TextFrame.Text);
            Assert.Contains("HELLO", shape.TextFrame.Text);
            Assert.Contains("hello", shape.TextFrame.Text);
        }
    }

    [SkippableFact]
    public void Execute_DefaultMatchCase_IsCaseInsensitive()
    {
        SkipIfNotWindows();
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

        if (!IsEvaluationMode())
        {
            var shape = pres.Slides[0].Shapes[0] as IAutoShape;
            Assert.NotNull(shape);
            Assert.Contains("new", shape.TextFrame.Text);
            Assert.DoesNotContain("TEST", shape.TextFrame.Text);
            Assert.DoesNotContain("test", shape.TextFrame.Text);
        }
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutFindText_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "replaceText", "New" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithoutReplaceText_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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
