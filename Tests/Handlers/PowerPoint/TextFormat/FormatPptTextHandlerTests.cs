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

    #region Table Shape Tests

    [Fact]
    public void Execute_WithTableShape_FormatsTableCells()
    {
        var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var colWidths = new double[] { 100, 100 };
        var rowHeights = new double[] { 30, 30 };
        var table = slide.Shapes.AddTable(100, 100, colWidths, rowHeights);
        table[0, 0].TextFrame.Text = "Cell A";
        table[1, 0].TextFrame.Text = "Cell B";
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", "Arial" },
            { "fontSize", 12.0 },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(presentation.Slides.Count > 0);
        Assert.True(slide.Shapes.Count > 0);
        AssertModified(context);
    }

    #endregion

    #region Basic Format Text Operations

    [Fact]
    public void Execute_FormatsAllSlides()
    {
        var presentation = CreatePresentationWithText("Sample text");
        var slideCount = presentation.Slides.Count;
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", "Arial" },
            { "fontSize", 14.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(slideCount, presentation.Slides.Count);
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

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(3, presentation.Slides.Count);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        Assert.True(presentation.Slides.Count > 0);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        Assert.True(presentation.Slides.Count > 0);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        Assert.True(presentation.Slides.Count > 0);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        Assert.True(presentation.Slides.Count > 0);
        AssertModified(context);
    }

    #endregion

    #region Alignment Tests

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("justify")]
    [InlineData("distributed")]
    public void Execute_WithAlignment_AppliesAlignment(string alignment)
    {
        var presentation = CreatePresentationWithText("Sample text");
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "alignment", alignment }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(presentation.Slides.Count > 0);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidAlignment_ThrowsArgumentException()
    {
        var presentation = CreatePresentationWithText("Sample text");
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "alignment", "invalid" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unsupported alignment", ex.Message);
    }

    #endregion

    #region Shape Indices Tests

    [Fact]
    public void Execute_WithShapeIndices_FormatsOnlyTargetShapes()
    {
        var presentation = CreatePresentationWithSlides(1);
        var slide = presentation.Slides[0];
        var shape0 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 50);
        shape0.TextFrame.Text = "Shape 0";
        var shape1 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 200, 50);
        shape1.TextFrame.Text = "Shape 1";
        var shapeCount = slide.Shapes.Count;
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", "[0]" },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(shapeCount, slide.Shapes.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var presentation = CreatePresentationWithText("Sample text");
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", "[99]" },
            { "bold", true }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shape index", ex.Message);
    }

    #endregion
}
