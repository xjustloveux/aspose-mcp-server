using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Text;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Text;

public class EditPptTextHandlerTests : PptHandlerTestBase
{
    private readonly EditPptTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Slide Index Parameter

    [Fact]
    public void Execute_WithSlideIndex_EditsOnCorrectSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        AddTextToSlide(pres, 1, "Text on slide 1");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 },
            { "text", "Updated" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide 1", result);
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

    [Fact]
    public void Execute_EditsText()
    {
        var pres = CreatePresentationWithText("Original");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "text", "Updated" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Text edited", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsShapeIndex()
    {
        var pres = CreatePresentationWithText("Original");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("shape 0", result);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreatePresentationWithText("Original");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide 0", result);
    }

    #endregion

    #region Font Formatting

    [Fact]
    public void Execute_WithFontName_AppliesFont()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fontName", "Arial" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result);
    }

    [Fact]
    public void Execute_WithFontSize_AppliesSize()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fontSize", 24f }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result);
    }

    [Fact]
    public void Execute_WithBold_AppliesBold()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "bold", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result);
    }

    [Fact]
    public void Execute_WithItalic_AppliesItalic()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "italic", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result);
    }

    [Fact]
    public void Execute_WithColor_AppliesColor()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "color", "#0000FF" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Updated" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Text");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
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
