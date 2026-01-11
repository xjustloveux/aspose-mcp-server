using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class GroupShapeDetailProviderTests : TestBase
{
    private readonly GroupShapeDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnGroup()
    {
        Assert.Equal("Group", _provider.TypeName);
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
    public void CanHandle_WithGroupShape_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var group = slide.Shapes.AddGroupShape();

        var result = _provider.CanHandle(group);

        Assert.True(result);
    }

    [Fact]
    public void GetDetails_WithGroupShape_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var group = slide.Shapes.AddGroupShape();

        var details = _provider.GetDetails(group, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithNonGroupShape_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.Null(details);
    }
}
