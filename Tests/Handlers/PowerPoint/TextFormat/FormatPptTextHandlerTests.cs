using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.TextFormat;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.TextFormat;

public class FormatPptTextHandlerTests : PptHandlerTestBase
{
    private readonly FormatPptTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Format()
    {
        Assert.Equal("format", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static void AddTextToSlides(Presentation presentation)
    {
        foreach (var slide in presentation.Slides)
        {
            var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
            shape.TextFrame.Text = "Test text";
        }
    }

    #endregion

    #region Basic Format Text Operations

    [Fact]
    public void Execute_FormatsAllSlides()
    {
        var presentation = CreatePresentationWithText("Sample text");
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", "Arial" },
            { "fontSize", 14.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("1 slides", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_FormatsSpecificSlides()
    {
        var presentation = CreatePresentationWithSlides(3);
        AddTextToSlides(presentation);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", "[0, 2]" },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("2 slides", result.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var presentation = CreatePresentationWithSlides(2);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndices", "[999]" },
            { "bold", true }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_AppliesBoldFormatting()
    {
        var presentation = CreatePresentationWithText("Sample text");
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("applied", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_AppliesItalicFormatting()
    {
        var presentation = CreatePresentationWithText("Sample text");
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "italic", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("applied", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_AppliesColorFormatting()
    {
        var presentation = CreatePresentationWithText("Sample text");
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "color", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("applied", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_AppliesMultipleFormattingOptions()
    {
        var presentation = CreatePresentationWithText("Sample text");
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", "Times New Roman" },
            { "fontSize", 16.0 },
            { "bold", true },
            { "italic", true },
            { "color", "#0000FF" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("applied", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
