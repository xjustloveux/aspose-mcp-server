using System.Text.Json;
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

    [Fact]
    public void GetDetails_WithGroupContainingShapes_ShouldReturnChildDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];

        var group = slide.Shapes.AddGroupShape();
        group.X = 10;
        group.Y = 10;
        group.Width = 220;
        group.Height = 100;

        var rect1 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        var rect2 = slide.Shapes.AddAutoShape(ShapeType.Ellipse, 120, 10, 100, 100);

        group.Shapes.AddClone(rect1);
        group.Shapes.AddClone(rect2);

        slide.Shapes.Remove(rect1);
        slide.Shapes.Remove(rect2);

        var details = _provider.GetDetails(group, presentation);
        Assert.NotNull(details);

        var json = JsonSerializer.Serialize(details);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.Equal(2, root.GetProperty("childCount").GetInt32());
        Assert.True(root.TryGetProperty("children", out var children));
        Assert.Equal(2, children.GetArrayLength());
    }

    [Fact]
    public void GetDetails_WithEmptyGroup_ShouldReturnNullChildren()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var group = slide.Shapes.AddGroupShape();

        var details = _provider.GetDetails(group, presentation);
        Assert.NotNull(details);

        var json = JsonSerializer.Serialize(details);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.Equal(0, root.GetProperty("childCount").GetInt32());
        Assert.Equal(JsonValueKind.Null, root.GetProperty("children").ValueKind);
    }
}
