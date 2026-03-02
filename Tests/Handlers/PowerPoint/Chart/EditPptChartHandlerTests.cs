using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Charts;
using AsposeMcpServer.Handlers.PowerPoint.Chart;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Chart;

[SupportedOSPlatform("windows")]
public class EditPptChartHandlerTests : PptHandlerTestBase
{
    private readonly EditPptChartHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Edit()
    {
        SkipIfNotWindows();
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [SkippableFact]
    public void Execute_EditsChartTitle()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "title", "Updated Title" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var chart = pres.Slides[0].Shapes.OfType<IChart>().First();
        Assert.True(chart.HasTitle);
        if (!IsEvaluationMode()) Assert.Equal("Updated Title", chart.ChartTitle.TextFrameForOverriding.Text);

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

    [SkippableFact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Chart Type Operations

    [SkippableFact]
    public void Execute_WithChartType_ChangesChartType()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "chartType", "Pie" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var chart = pres.Slides[0].Shapes.OfType<IChart>().First();
        Assert.Equal(ChartType.Pie, chart.Type);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithTitleAndChartType_UpdatesBoth()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "title", "New Title" },
            { "chartType", "Line" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var chart = pres.Slides[0].Shapes.OfType<IChart>().First();
        Assert.Equal(ChartType.Line, chart.Type);
        Assert.True(chart.HasTitle);
        if (!IsEvaluationMode()) Assert.Equal("New Title", chart.ChartTitle.TextFrameForOverriding.Text);

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithNoUpdates_StillReturnsSuccess()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithChart();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var chart = pres.Slides[0].Shapes.OfType<IChart>().First();
        Assert.Equal(ChartType.ClusteredColumn, chart.Type);
    }

    #endregion
}
