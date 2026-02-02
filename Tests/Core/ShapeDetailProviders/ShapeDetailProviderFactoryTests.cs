using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;
using AsposeMcpServer.Core.ShapeDetailProviders.Providers;
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
        Assert.Equal("AutoShape", provider.TypeName);
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
        Assert.Equal("Table", provider.TypeName);
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
        Assert.Equal("Connector", provider.TypeName);
    }

    #endregion

    #region GetShapeDetails Tests

    [SkippableFact]
    public void GetShapeDetails_WithAutoShape_ShouldReturnDetails()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Text content is truncated in evaluation mode");
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.TextFrame.Text = "Test Text";

        var (typeName, details) = ShapeDetailProviderFactory.GetShapeDetails(shape, presentation);

        Assert.Equal("AutoShape", typeName);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.Equal("Rectangle", autoShape.ShapeType);
        Assert.Equal("Test Text", autoShape.Text);
        Assert.True(autoShape.HasTextFrame);
        Assert.True(autoShape.ParagraphCount >= 1);
    }

    [Fact]
    public void GetShapeDetails_WithTable_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var (typeName, details) = ShapeDetailProviderFactory.GetShapeDetails(table, presentation);

        Assert.Equal("Table", typeName);
        var tableDetails = Assert.IsType<TableDetails>(details);

        Assert.Equal(2, tableDetails.Rows);
        Assert.Equal(2, tableDetails.Columns);
        Assert.Equal(4, tableDetails.TotalCells);
    }

    #endregion
}
