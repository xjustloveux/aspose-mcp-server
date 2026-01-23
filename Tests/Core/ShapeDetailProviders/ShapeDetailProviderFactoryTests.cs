using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class ShapeDetailProviderFactoryTests : TestBase
{
    #region GetProvider Tests

    [Fact]
    public void GetProvider_WithAutoShape_ShouldReturnAutoShapeProvider()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var provider = ShapeDetailProviderFactory.GetProvider(shape);

        Assert.NotNull(provider);
        Assert.IsType<AutoShapeDetailProvider>(provider);
    }

    [Fact]
    public void GetProvider_WithTable_ShouldReturnTableProvider()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var provider = ShapeDetailProviderFactory.GetProvider(table);

        Assert.NotNull(provider);
        Assert.IsType<TableDetailProvider>(provider);
    }

    [Fact]
    public void GetProvider_WithConnector_ShouldReturnConnectorProvider()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var connector = slide.Shapes.AddConnector(ShapeType.StraightConnector1, 10, 10, 100, 100);

        var provider = ShapeDetailProviderFactory.GetProvider(connector);

        Assert.NotNull(provider);
        Assert.IsType<ConnectorDetailProvider>(provider);
    }

    #endregion

    #region GetShapeDetails Tests

    [Fact]
    public void GetShapeDetails_WithAutoShape_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.TextFrame.Text = "Test Text";

        var (typeName, properties) = ShapeDetailProviderFactory.GetShapeDetails(shape, presentation);

        Assert.Equal("AutoShape", typeName);
        Assert.NotNull(properties);
    }

    [Fact]
    public void GetShapeDetails_WithTable_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var (typeName, properties) = ShapeDetailProviderFactory.GetShapeDetails(table, presentation);

        Assert.Equal("Table", typeName);
        Assert.NotNull(properties);
    }

    #endregion
}
