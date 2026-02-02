using Aspose.Slides.Charts;
using AsposeMcpServer.Handlers.PowerPoint.Chart;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Chart;

public class AddPptChartHandlerTests : PptHandlerTestBase
{
    private readonly AddPptChartHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsChart()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "chartType", "Column" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var slide = pres.Slides[0];
        var chart = slide.Shapes.OfType<IChart>().FirstOrDefault();
        Assert.NotNull(chart);
        Assert.Equal(ChartType.ClusteredColumn, chart.Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithTitle_AddsChartWithTitle()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "chartType", "Pie" },
            { "title", "My Chart" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var slide = pres.Slides[0];
        var chart = slide.Shapes.OfType<IChart>().FirstOrDefault();
        Assert.NotNull(chart);
        Assert.Equal(ChartType.Pie, chart.Type);
        Assert.True(chart.HasTitle);
        if (!IsEvaluationMode()) Assert.Equal("My Chart", chart.ChartTitle.TextFrameForOverriding.Text);

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomPosition_AddsChartAtPosition()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "chartType", "Line" },
            { "x", 100f },
            { "y", 100f },
            { "width", 400f },
            { "height", 300f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var slide = pres.Slides[0];
        var chart = slide.Shapes.OfType<IChart>().FirstOrDefault();
        Assert.NotNull(chart);
        Assert.Equal(ChartType.Line, chart.Type);
        Assert.Equal(100f, chart.X);
        Assert.Equal(100f, chart.Y);
        Assert.Equal(400f, chart.Width);
        Assert.Equal(300f, chart.Height);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "chartType", "Column" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutChartType_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
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
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "chartType", "Column" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
