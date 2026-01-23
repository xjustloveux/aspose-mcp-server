using Aspose.Slides;
using Aspose.Slides.Charts;
using AsposeMcpServer.Handlers.PowerPoint.Chart;
using AsposeMcpServer.Results.PowerPoint.Chart;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Chart;

public class GetPptChartDataHandlerTests : PptHandlerTestBase
{
    private readonly GetPptChartDataHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetData()
    {
        Assert.Equal("get_data", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithChart()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 400);
        return pres;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsChartData()
    {
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartDataPptResult>(res);

        Assert.NotNull(result.ChartType);
        Assert.NotNull(result.Categories);
        Assert.NotNull(result.Series);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetChartDataPptResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.ChartType);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
