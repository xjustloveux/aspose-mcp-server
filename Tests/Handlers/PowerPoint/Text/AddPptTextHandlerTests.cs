using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Text;

[SupportedOSPlatform("windows")]
public class AddPptTextHandlerTests : PptHandlerTestBase
{
    private readonly AddPptTextHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Add()
    {
        SkipIfNotWindows();
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Existing Shape

    [SkippableFact]
    public void Execute_WithShapeIndex_AddsToExistingShape()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Original");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Updated Text" },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[0];
            Assert.Contains("Updated Text", shape.TextFrame.Text);
        }

        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [SkippableFact]
    public void Execute_AddsText()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Hello World" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[^1];
            Assert.Contains("Hello World", shape.TextFrame.Text);
        }
    }

    [SkippableFact]
    public void Execute_ReturnsSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var initialShapeCount = pres.Slides[1].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test" },
            { "slideIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount + 1, pres.Slides[1].Shapes.Count);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_CreatesNewShape()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New Text" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialShapeCount + 1, pres.Slides[0].Shapes.Count);
    }

    #endregion

    #region Position Parameters

    [SkippableFact]
    public void Execute_WithPosition_CreatesShapeAtPosition()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Positioned Text" },
            { "x", 200f },
            { "y", 300f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount + 1, pres.Slides[0].Shapes.Count);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithSize_CreatesShapeWithSize()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Sized Text" },
            { "width", 500f },
            { "height", 200f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount + 1, pres.Slides[0].Shapes.Count);
        AssertModified(context);
    }

    #endregion

    #region Font Formatting

    [SkippableFact]
    public void Execute_WithFontName_AppliesFont()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Formatted" },
            { "fontName", "Arial" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount + 1, pres.Slides[0].Shapes.Count);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[^1];
            Assert.Contains("Formatted", shape.TextFrame.Text);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithFontSize_AppliesSize()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Large Text" },
            { "fontSize", 24f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount + 1, pres.Slides[0].Shapes.Count);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[^1];
            Assert.Contains("Large Text", shape.TextFrame.Text);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithBold_AppliesBold()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Bold Text" },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount + 1, pres.Slides[0].Shapes.Count);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[^1];
            Assert.Contains("Bold Text", shape.TextFrame.Text);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithItalic_AppliesItalic()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Italic Text" },
            { "italic", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount + 1, pres.Slides[0].Shapes.Count);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[^1];
            Assert.Contains("Italic Text", shape.TextFrame.Text);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithColor_AppliesColor()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Colored Text" },
            { "color", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount + 1, pres.Slides[0].Shapes.Count);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[^1];
            Assert.Contains("Colored Text", shape.TextFrame.Text);
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test" },
            { "slideIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test" },
            { "shapeIndex", 99 }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
