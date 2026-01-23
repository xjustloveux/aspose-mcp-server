using Aspose.Slides;
using Aspose.Slides.Charts;
using AsposeMcpServer.Handlers.PowerPoint.Chart;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Chart;

public class DeletePptChartHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptChartHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesChart()
    {
        var pres = CreatePresentationWithChart();
        var initialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(initialCount - 1, pres.Slides[0].Shapes.Count);
        AssertModified(context);
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
