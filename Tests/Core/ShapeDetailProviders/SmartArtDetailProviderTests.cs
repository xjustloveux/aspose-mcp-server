using Aspose.Slides;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class SmartArtDetailProviderTests : TestBase
{
    private readonly SmartArtDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnSmartArt()
    {
        Assert.Equal("SmartArt", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }

    [Fact]
    public void CanHandle_WithSmartArt_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var smartArt = slide.Shapes.AddSmartArt(10, 10, 400, 300, SmartArtLayoutType.BasicBlockList);

        var result = _provider.CanHandle(smartArt);

        Assert.True(result);
    }

    [Fact]
    public void GetDetails_WithSmartArt_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var smartArt = slide.Shapes.AddSmartArt(10, 10, 400, 300, SmartArtLayoutType.BasicBlockList);

        var details = _provider.GetDetails(smartArt, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithNonSmartArt_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.Null(details);
    }
}
