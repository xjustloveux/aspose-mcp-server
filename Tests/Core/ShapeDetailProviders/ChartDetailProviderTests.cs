using Aspose.Slides;
using Aspose.Slides.Charts;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;
using AsposeMcpServer.Core.ShapeDetailProviders.Providers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class ChartDetailProviderTests : TestBase
{
    private readonly ChartDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnChart()
    {
        Assert.Equal("Chart", _provider.TypeName);
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
    public void CanHandle_WithChart_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 10, 10, 400, 300);

        var result = _provider.CanHandle(chart);

        Assert.True(result);
    }

    [Fact]
    public void GetDetails_WithChart_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 10, 10, 400, 300);

        var details = _provider.GetDetails(chart, presentation);
        var chartDetails = Assert.IsType<ChartDetails>(details);

        Assert.Equal("ClusteredColumn", chartDetails.ChartType);
        Assert.True(chartDetails.SeriesCount >= 0);
        Assert.True(chartDetails.CategoryCount >= 0);
    }

    [Fact]
    public void GetDetails_WithNonChart_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.Null(details);
    }

    [Fact]
    public void GetDetails_WithLegend_ShouldReturnLegendPosition()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 10, 10, 400, 300);
        chart.HasLegend = true;

        var details = _provider.GetDetails(chart, presentation);
        var chartDetails = Assert.IsType<ChartDetails>(details);

        Assert.True(chartDetails.HasLegend);
        Assert.NotNull(chartDetails.LegendPosition);
    }

    [Fact]
    public void GetDetails_WithoutLegend_ShouldReturnNullLegendPosition()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 10, 10, 400, 300);
        chart.HasLegend = false;

        var details = _provider.GetDetails(chart, presentation);
        var chartDetails = Assert.IsType<ChartDetails>(details);

        Assert.False(chartDetails.HasLegend);
        Assert.Null(chartDetails.LegendPosition);
    }
}
