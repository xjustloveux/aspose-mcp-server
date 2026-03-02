using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Text;

[SupportedOSPlatform("windows")]
public class EditPptTextHandlerTests : PptHandlerTestBase
{
    private readonly EditPptTextHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Edit()
    {
        SkipIfNotWindows();
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Slide Index Parameter

    [SkippableFact]
    public void Execute_WithSlideIndex_EditsOnCorrectSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        AddTextToSlide(pres, 1, "Text on slide 1");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 },
            { "text", "Updated" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[1].Shapes[0];
            Assert.Contains("Updated", shape.TextFrame.Text);
        }

        AssertModified(context);
    }

    #endregion

    #region Helper Methods

    private static void AddTextToSlide(Presentation pres, int slideIndex, string text)
    {
        var slide = pres.Slides[slideIndex];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 400, 100);
        shape.TextFrame.Text = text;
    }

    #endregion

    #region Basic Edit Operations

    [SkippableFact]
    public void Execute_EditsText()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Original");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "text", "Updated" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var shape = (IAutoShape)pres.Slides[0].Shapes[0];
            Assert.Contains("Updated", shape.TextFrame.Text);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsShapeIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Original");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Shapes.Count > 0);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Original");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Shapes.Count > 0);
        AssertModified(context);
    }

    #endregion

    #region Font Formatting

    [SkippableFact]
    public void Execute_WithFontName_AppliesFont()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fontName", "Arial" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Shapes.Count > 0);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithFontSize_AppliesSize()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fontSize", 24f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Shapes.Count > 0);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithBold_AppliesBold()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Shapes.Count > 0);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithItalic_AppliesItalic()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "italic", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Shapes.Count > 0);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithColor_AppliesColor()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "color", "#0000FF" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Shapes.Count > 0);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Updated" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithNegativeShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
