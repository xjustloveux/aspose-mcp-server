using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class AutoShapeDetailProviderTests : TestBase
{
    private readonly AutoShapeDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnAutoShape()
    {
        Assert.Equal("AutoShape", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.True(result);
    }

    [Fact]
    public void CanHandle_WithTable_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50], [30]);

        var result = _provider.CanHandle(table);

        Assert.False(result);
    }

    [Fact]
    public void GetDetails_WithAutoShape_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.TextFrame.Text = "Sample Text";

        var details = _provider.GetDetails(shape, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithNonAutoShape_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50], [30]);

        var details = _provider.GetDetails(table, presentation);

        Assert.Null(details);
    }

    [Fact]
    public void GetDetails_WithEmptyTextFrame_ShouldIncludeNullText()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.NotNull(details);
    }
}
