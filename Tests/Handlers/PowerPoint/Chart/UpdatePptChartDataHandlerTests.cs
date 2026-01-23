using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Charts;
using AsposeMcpServer.Handlers.PowerPoint.Chart;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Chart;

public class UpdatePptChartDataHandlerTests : PptHandlerTestBase
{
    private readonly UpdatePptChartDataHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_UpdateData()
    {
        Assert.Equal("update_data", _handler.Operation);
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

    #region Basic Update Operations

    [Fact]
    public void Execute_UpdatesChartData()
    {
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var data = new JsonObject
        {
            ["categories"] = new JsonArray("Q1", "Q2", "Q3"),
            ["series"] = new JsonArray(
                new JsonObject
                {
                    ["name"] = "Sales",
                    ["values"] = new JsonArray(100.0, 200.0, 150.0)
                }
            )
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "data", data }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithClearExisting_ClearsAndUpdates()
    {
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "clearExisting", true }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoChanges_ReturnsNoChangesMessage()
    {
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("no changes", result.Message, StringComparison.OrdinalIgnoreCase);
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
