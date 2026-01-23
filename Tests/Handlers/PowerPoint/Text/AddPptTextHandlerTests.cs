using AsposeMcpServer.Handlers.PowerPoint.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Text;

public class AddPptTextHandlerTests : PptHandlerTestBase
{
    private readonly AddPptTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Existing Shape

    [Fact]
    public void Execute_WithShapeIndex_AddsToExistingShape()
    {
        var pres = CreatePresentationWithText("Original");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Updated Text" },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Text added", result.Message);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsText()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Hello World" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Text added", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test" },
            { "slideIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("slide 1", result.Message);
    }

    [Fact]
    public void Execute_CreatesNewShape()
    {
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

    [Fact]
    public void Execute_WithPosition_CreatesShapeAtPosition()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Positioned Text" },
            { "x", 200f },
            { "y", 300f }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Text added", result.Message);
    }

    [Fact]
    public void Execute_WithSize_CreatesShapeWithSize()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Sized Text" },
            { "width", 500f },
            { "height", 200f }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Text added", result.Message);
    }

    #endregion

    #region Font Formatting

    [Fact]
    public void Execute_WithFontName_AppliesFont()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Formatted" },
            { "fontName", "Arial" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Text added", result.Message);
    }

    [Fact]
    public void Execute_WithFontSize_AppliesSize()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Large Text" },
            { "fontSize", 24f }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Text added", result.Message);
    }

    [Fact]
    public void Execute_WithBold_AppliesBold()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Bold Text" },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Text added", result.Message);
    }

    [Fact]
    public void Execute_WithItalic_AppliesItalic()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Italic Text" },
            { "italic", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Text added", result.Message);
    }

    [Fact]
    public void Execute_WithColor_AppliesColor()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Colored Text" },
            { "color", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Text added", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test" },
            { "slideIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
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
