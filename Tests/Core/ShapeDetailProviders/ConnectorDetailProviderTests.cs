using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;
using AsposeMcpServer.Core.ShapeDetailProviders.Providers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class ConnectorDetailProviderTests : TestBase
{
    private readonly ConnectorDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnConnector()
    {
        Assert.Equal("Connector", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithConnector_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var connector = slide.Shapes.AddConnector(ShapeType.StraightConnector1, 10, 10, 100, 100);

        var result = _provider.CanHandle(connector);

        Assert.True(result);
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
    public void GetDetails_WithConnector_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var connector = slide.Shapes.AddConnector(ShapeType.StraightConnector1, 10, 10, 100, 100);

        var details = _provider.GetDetails(connector, presentation);
        var connectorDetails = Assert.IsType<ConnectorDetails>(details);

        Assert.Equal("StraightConnector1", connectorDetails.ConnectorType);
        Assert.Null(connectorDetails.StartShapeConnectedTo);
        Assert.Null(connectorDetails.EndShapeConnectedTo);
    }

    [Fact]
    public void GetDetails_WithNonConnector_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.Null(details);
    }

    [Fact]
    public void GetDetails_WithLineFormat_ShouldReturnLineProperties()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var connector = slide.Shapes.AddConnector(ShapeType.StraightConnector1, 10, 10, 100, 100);
        connector.LineFormat.FillFormat.FillType = FillType.Solid;
        connector.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;
        connector.LineFormat.Width = 2.5;
        connector.LineFormat.DashStyle = LineDashStyle.DashDot;

        var details = _provider.GetDetails(connector, presentation);
        var connectorDetails = Assert.IsType<ConnectorDetails>(details);

        Assert.Equal("#0000FF", connectorDetails.LineColor);
        Assert.Equal(2.5, connectorDetails.LineWidth);
        Assert.Equal("DashDot", connectorDetails.LineDashStyle);
    }

    [Fact]
    public void GetDetails_WithDefaultConnector_ShouldNotFailOnLineFormat()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var connector = slide.Shapes.AddConnector(ShapeType.StraightConnector1, 10, 10, 100, 100);

        var details = _provider.GetDetails(connector, presentation);
        var connectorDetails = Assert.IsType<ConnectorDetails>(details);

        Assert.NotNull(connectorDetails);
    }
}
